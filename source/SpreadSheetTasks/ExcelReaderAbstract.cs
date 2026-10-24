using System;
using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace SpreadSheetTasks
{
    //some code from https://github.com/MarkPflug/Sylvan.Data.Excel
    //some code from https://github.com/ExcelDataReader/ExcelDataReader
    public abstract class ExcelReaderAbstract
    {
        protected static readonly CultureInfo invariantCultureInfo = CultureInfo.InvariantCulture;

        /// <summary>
        /// Number of fields (columns) in the current row.
        /// </summary>
        public int FieldCount { get; set; }

        /// <summary>
        /// Number of sheets in the workbook.
        /// </summary>
        public virtual int ResultsCount { get; }

        /// <summary>
        /// Name of the sheet currently being read. Set this property to switch sheets.
        /// </summary>
        public virtual string ActualSheetName { get; set; }

        /// <summary>
        /// Number of data rows in the current sheet (excluding the header row).
        /// Returns <c>-1</c> when the row count is not available from the file metadata.
        /// </summary>
        public virtual int RowCount { get => -1; }

        /// <summary>
        /// Opens an Excel file for reading. The file format (xlsx/xlsb) is detected from the file extension.
        /// </summary>
        /// <param name="path">Path to the .xlsx or .xlsb file.</param>
        /// <param name="readSharedStrings">If true, load shared strings into memory for string-typed cells.</param>
        /// <param name="updateMode">If true, open the file in update mode so the file can be modified (e.g. via <see cref="XlsxOrXlsbReadOrEdit.ReplaceSheetData"/>).</param>
        /// <param name="encoding">Optional encoding hint; not used for native .xlsx/.xlsb files.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="path"/> is null.</exception>
        /// <exception cref="ArgumentException">Thrown when <paramref name="path"/> is empty or whitespace.</exception>
        /// <exception cref="FileNotFoundException">Thrown when the file does not exist.</exception>
        public abstract void Open(string path, bool readSharedStrings = true, bool updateMode = false, Encoding? encoding = null);

        /// <summary>
        /// Reads the next row from the current sheet.
        /// </summary>
        /// <returns>true if a row was read; false when the sheet has no more rows.</returns>
        public abstract bool Read();

        /// <summary>
        /// Approximate read position (0-100) of the current row in the sheet stream, useful for progress reporting.
        /// </summary>
        public virtual double RelativePositionInStream() => 50.0;

        /// <summary>
        /// use only after read first row  = GetValue + ToString
        /// </summary>
        /// <param name="i">column number</param>
        /// <returns></returns>
        public string GetName(int i)
        {
            return GetValue(i).ToString();
        }

        //public virtual string GetName(int i) => dbReader.GetName(i);
        public Type GetFieldType(int i)
        {
            return innerRow[i].type switch
            {
                ExcelDataType.Null => typeof(DBNull),
                ExcelDataType.Int64 => typeof(Int64),
                ExcelDataType.Double => typeof(Double),
                ExcelDataType.DateTime => typeof(DateTime),
                ExcelDataType.Boolean => typeof(bool),
                _ => typeof(string),
            };
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public object GetValue(int i)
        {
            ref var w = ref innerRow[i];
            return w.type switch
            {
                ExcelDataType.Null => DBNull.Value,
                ExcelDataType.Int32 => w.int32Value,
                ExcelDataType.Int64 => w.int64Value,
                ExcelDataType.Double => w.doubleValue,
                ExcelDataType.DateTime => w.dtValue,
                ExcelDataType.Boolean => w.boolValue,
                ExcelDataType.String => w.strValue,
                //case ExcelDataType.Error:
                //    return "error in cell";
                _ => null,
            };
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public void GetValues(object[] row)
        {
            for (int i = 0; i < row.Length; i++)
            {
                ref var w = ref innerRow[i];
                row[i] = w.type switch
                {
                    ExcelDataType.Null => DBNull.Value,
                    ExcelDataType.Int32 => w.int32Value,
                    ExcelDataType.Int64 => w.int64Value,
                    ExcelDataType.Double => w.doubleValue,
                    ExcelDataType.DateTime => w.dtValue,
                    ExcelDataType.Boolean => w.boolValue,
                    ExcelDataType.String => w.strValue,
                    _ => null,
                };
            }
        }


        public ref FieldInfo GetNativeValue(int i)
        {
            return ref innerRow![i];
        }

        public ref FieldInfo[] GetNativeValues()
        {
            return ref innerRow!;
        }

        /// <summary>
        /// The current row as a span of <see cref="FieldInfo"/> values, without boxing numbers,
        /// dates or booleans into an <c>object[]</c>. The span is only valid until the next
        /// <see cref="Read"/> call.
        /// </summary>
        public ReadOnlySpan<FieldInfo> CurrentRow
        {
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            get
            {
                if (innerRow is null)
                    return default;
                int count = Math.Min(FieldCount, innerRow.Length);
                return innerRow.AsSpan(0, count);
            }
        }

        /// <summary>
        /// Returns the Excel cell-type for the cell at <paramref name="i"/> in the current row.
        /// Use this when you want to know the underlying cell type (e.g. <see cref="ExcelDataType.Int64"/>,
        /// <see cref="ExcelDataType.Double"/>, <see cref="ExcelDataType.DateTime"/>,
        /// <see cref="ExcelDataType.String"/>, <see cref="ExcelDataType.Boolean"/>,
        /// <see cref="ExcelDataType.Null"/>) without touching the raw native value.
        /// For most callers, the typed getters (<see cref="GetInt32"/>, <see cref="GetDouble"/>,
        /// <see cref="GetDateTime"/>, <see cref="GetString"/>) or <see cref="GetValue(int)"/> are simpler.
        /// </summary>
        /// <param name="i">Zero-based column index.</param>
        /// <returns>The <see cref="ExcelDataType"/> for the cell.</returns>
        public ExcelDataType GetExcelDataType(int i)
        {
            return innerRow[i].type;
        }

        /// <summary>
        /// Returns the Excel cell type name and the raw native value for the cell at <paramref name="i"/>.
        /// The native value is a <see cref="ref FieldInfo"/> that exposes all storage fields at once
        /// (typed via <see cref="ExcelDataType"/>). This is the same as calling
        /// <see cref="GetExcelDataType(int)"/> plus <see cref="GetNativeValue(int)"/>, returned together
        /// for convenience.
        /// </summary>
        /// <param name="i">Zero-based column index.</param>
        /// <param name="cellType">The <see cref="ExcelDataType"/> name for the cell.</param>
        /// <param name="nativeValue">A <see cref="ref FieldInfo"/> reference into the reader's row buffer.</param>
        public void GetExcelDataType(int i, out string cellType, out object? nativeValue)
        {
            ref var w = ref innerRow[i];
            cellType = w.type.ToString();
            nativeValue = w.type switch
            {
                ExcelDataType.Null => null,
                ExcelDataType.Int32 => w.int32Value,
                ExcelDataType.Int64 => w.int64Value,
                ExcelDataType.Double => w.doubleValue,
                ExcelDataType.DateTime => w.dtValue,
                ExcelDataType.Boolean => w.boolValue,
                ExcelDataType.String => w.strValue,
                _ => null,
            };
        }

        public bool TreatAllColumnsAsText { get; set; } = false;
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public virtual string GetString(int i)
        {
            ref var w = ref innerRow[i];
            if (w.type == ExcelDataType.String)
            {
                return w.strValue ?? string.Empty;
            }
            if (w.type == ExcelDataType.Null)
            {
                return string.Empty;
            }
            return w.type switch
            {
                ExcelDataType.Int32 => w.int32Value.ToString(invariantCultureInfo),
                ExcelDataType.Int64 => w.int64Value.ToString(invariantCultureInfo),
                ExcelDataType.Double => w.doubleValue.ToString(invariantCultureInfo),
                ExcelDataType.DateTime => w.dtValue.ToString(invariantCultureInfo),
                ExcelDataType.Boolean => w.boolValue ? bool.TrueString : bool.FalseString,
                _ => GetValue(i)?.ToString() ?? string.Empty,
            };
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public DateTime GetDateTime(int i)
        {
            ref var w = ref innerRow[i];
            if (w.type == ExcelDataType.DateTime)
            {
                return w.dtValue;
            }
            else
            {
                throw new InvalidCastException();
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public Int64 GetInt64(int i)
        {
            ref var w = ref innerRow[i];
            switch (w.type)
            {
                case ExcelDataType.Int64:
                    return w.int64Value;
                case ExcelDataType.Int32:
                    return w.int32Value;
                case ExcelDataType.Double:
                    return Convert.ToInt64(w.doubleValue);
                default:
                    throw new InvalidCastException();
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public Int32 GetInt32(int i)
        {
            ref var w = ref innerRow[i];
            switch (w.type)
            {
                case ExcelDataType.Int32:
                    return w.int32Value;
                case ExcelDataType.Int64:
                    return checked((int)w.int64Value);
                case ExcelDataType.Double:
                    return Convert.ToInt32(w.doubleValue);
                default:
                    throw new InvalidCastException();
            }
        }


        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public double GetDouble(int i)
        {
            ref var w = ref innerRow[i];
            switch (w.type)
            {
                case ExcelDataType.Double:
                    return w.doubleValue;
                case ExcelDataType.Int64:
                    return w.int64Value;
                case ExcelDataType.Int32:
                    return w.int32Value;
                default:
                    throw new InvalidCastException();
            }
        }

        /// <summary>
        /// Returns the names of all sheets in the workbook, in the order they appear in the file.
        /// </summary>
        public abstract string[] GetSheetNames();

        /// <summary>
        /// Obsolete. Use <see cref="GetSheetNames"/> instead.
        /// </summary>
        [Obsolete("Use GetSheetNames() instead. This misspelled member will be removed in a future release.")]
        public string[] GetScheetNames() => GetSheetNames();

        protected FieldInfo[] innerRow;

        public abstract void Dispose();
        //public virtual void Dispose()
        //{
        //    dbReader.Dispose();
        //}

        //Sylvan

    }

    [StructLayout(LayoutKind.Explicit)]
    public struct FieldInfo
    {
        [FieldOffset(8)]
        public ExcelDataType type;
        [FieldOffset(12)]
        public Int64 int64Value;
        [FieldOffset(12)]
        public Int32 int32Value;
        [FieldOffset(12)]
        public bool boolValue;
        [FieldOffset(12)]
        public double doubleValue;
        [FieldOffset(12)]
        public DateTime dtValue;
        [FieldOffset(0)]
        public string strValue;
    }

    public enum ExcelDataType
    {
        /// <summary>
        /// A cell that contains no value.
        /// </summary>
        Null = 0,
        /// <summary>
        /// Number
        /// </summary>
        Int32,
        Int64,
        Double,
        /// <summary>
        /// A DateTime value. This is an uncommonly used representation in .xlsx files.
        /// </summary>
        DateTime,
        /// <summary>
        /// A text field.
        /// </summary>
        String,
        /// <summary>
        /// A formula cell that contains a boolean.
        /// </summary>
        Boolean,
        /// <summary>
        /// A formula cell that contains an error.
        /// </summary>
        Error,
    }

}
