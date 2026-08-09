using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;

//some code from https://github.com/MarkPflug/Sylvan.Data.Excel
//some code from https://github.com/ExcelDataReader/ExcelDataReader

namespace SpreadSheetTasks;

public sealed class XlsxOrXlsbReadOrEdit : ExcelReaderAbstract, IDisposable
{
    private ZipArchive _xlsxArchive;
    private readonly Dictionary<string, string> _worksheetIdToName = [];
    private readonly Dictionary<string, string> _worksheetNameToId = [];
    private readonly Dictionary<string, string> _worksheetIdToLocation = [];
    private Dictionary<int, string> _pivotCacheIdToRid;
    private Dictionary<string, string> _pivotCacheRidToLocation;
    //private Dictionary<int, string> _worksheetSheetIdToId = new Dictionary<int, string>();

    private string[] _sharedStringArray;
    private StyleInfo[] _stylesCellXfsArray;

    private readonly Dictionary<int, Type> _numberFormatsTypeDictionary = new Dictionary<int, Type>()
    {
        // to do
        {0,typeof(string)},
        {1,typeof(decimal?)},
        {2,typeof(decimal?)},
        {3,typeof(decimal?)},
        {4,typeof(decimal?)},
        {5,typeof(decimal?)},
        {6,typeof(decimal?)},
        {7,typeof(decimal?)},

        {9, typeof(decimal?)}, // 0%

        { 14,typeof(DateTime?)},
        { 15,typeof(DateTime?)},
        { 16,typeof(DateTime?)},
        { 17,typeof(DateTime?)},
        { 18,typeof(DateTime?)},
        { 19,typeof(DateTime?)},
        { 20,typeof(DateTime?)},
        { 21,typeof(DateTime?)},
        { 22,typeof(DateTime?)},

        { 44,typeof(decimal?)}
    };

    internal readonly static HashSet<string> _dateExcelMasks = new HashSet<string>()
    {
        @"[$-F800]dddd\,\ mmmm\ dd\,\ yyyy",
        @"d\-mm;@",
        @"yy\-mm\-dd;@",
        @"[$-415]d\ mmm;@",
        @"[$-415]d\ mmm\ yy;@",
        @"[$-415]dd\ mmm\ yy;@",
        @"[$-415]mmm\ yy;@",
        @"[$-415]mmmm\ yy;@",
        @"[$-415]d\ mmmm\ yyyy;@",
        @"yyyy\-mm\-dd\ hh:mm",
        @"yyyy\-mm\-dd\ hh:mm:ss",
        @"yyyy\-mm\-dd;@",
        @"[$-409]dd\-mm\-yy\ h:mm\ AM/PM;@",
        @"dd\-mm\-yy\ h:mm;@",
        @"[$-415]mmmmm;@",
        @"[$-415]mmmmm\.yy;@",
        @"\-m\-yyyy;@",
        @"[$-415]d\-mmm\-yyyy;@",
        @"d\-m\-yyyy;@"
    };

    private string _sharedStringsLocation = null;
    private string _stylesLocation = null;
    //private string _themeLocation = null;
    private int _uniqueStringCount = -1;
    //private int _stringCount = -1;

    private Modes _mode = Modes.xlsx;
    enum Modes
    {
        xlsx, xlsb
    }

    private static readonly string _openXmlInfoString = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static readonly Dictionary<string, int> _letterToColumnNum = new Dictionary<string, int>();
    static XlsxOrXlsbReadOrEdit()
    {
        for (int i = 0; i < XlsxWriter._letters.Length; i++)
        {
            _letterToColumnNum[XlsxWriter._letters[i]] = i;
        }
    }
    public override void Dispose()
    {
        _xmlReader?.Dispose();
        _sheetStream?.Dispose();
        _biffReader?.Dispose();
        _xlsxArchive?.Dispose();
    }

    private void OpenXlsx(string path, bool readSharedStrings = true, bool updateMode = false)
    {
        _mode = Modes.xlsx;
        if (updateMode)
        {
            _xlsxArchive = new ZipArchive(new FileStream(path, FileMode.Open), ZipArchiveMode.Read | ZipArchiveMode.Update);
        }
        else
        {
            _xlsxArchive = new ZipArchive(new FileStream(path, FileMode.Open), ZipArchiveMode.Read);
        }

        var e1 = _xlsxArchive.GetEntry("xl/workbook.xml");
        using (var str = e1.Open())
        {
            using var reader = XmlReader.Create(str);
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Whitespace)
                {
                    continue;
                }

                if (reader.Name == "sheet")
                {
                    string name = reader.GetAttribute("name");
                    string rId = reader.GetAttribute("r:id");
                    _worksheetIdToName[rId] = name;
                    _worksheetNameToId[name] = rId;
                }
                else if (reader.Name == "pivotCache")
                {
                    _pivotCacheIdToRid ??= new Dictionary<int, string>();

                    if (!int.TryParse(reader.GetAttribute("cacheId"), out int cacheId))
                    {
                        throw new Exception("getting pivot table cache error");
                    }

                    string rId = reader.GetAttribute("r:id");
                    _pivotCacheIdToRid[cacheId] = rId;
                }
            }
        }

        _resultCount = _worksheetIdToName.Count;
        FillRels("xml");

        if (_stylesLocation != null)
        {
            FillStyles();
        }
        if (readSharedStrings && _sharedStringsLocation != null)
        {
            FillSharedStrings();
        }
    }

    private void OpenXlsb(string path, bool readSharedStrings = true)
    {
        _mode = Modes.xlsb;
        _xlsxArchive = new ZipArchive(new FileStream(path, FileMode.Open), ZipArchiveMode.Read);

        var e1 = _xlsxArchive.GetEntry("xl/workbook.bin");

        BiffReaderWriter reader;
        Stream? str = null;
        if (UseMemoryStreamInXlsb)
        {
            reader = new BiffReaderWriter(ReadEntryToBytes(e1.Open(), e1.Length));
        }
        else
        {
            str = new BufferedStream(e1.Open());
            reader = new BiffReaderWriter(str);
        }

        try
        {
            while (reader.ReadWorkbook())
            {
                if (reader._isSheet == true)
                {
                    string name = reader._workbookName;
                    string rId = reader._recId;
                    _worksheetIdToName[rId] = name;
                    _worksheetNameToId[name] = rId;
                }
            }
        }
        finally
        {
            str?.Dispose();
        }


        _resultCount = _worksheetIdToName.Count;

        FillRels("bin");


        if (_stylesLocation != null)
        {
            FillBinStyles();
        }

        if (readSharedStrings && _sharedStringsLocation != null)
        {
            FillBinSharedStrings();
        }

    }
    /// <summary>
    /// Initialize file to read (detection xlsx/xlsb by extension)
    /// </summary>
    /// <param name="path"></param>
    /// <param name="readSharedStrings"></param>
    /// <param name="updateMode"></param>
    public override void Open(string path, bool readSharedStrings = true, bool updateMode = false, Encoding? encoding = null)
    {
        ArgumentNullException.ThrowIfNull(path);
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        if (path.EndsWith("xlsb", StringComparison.OrdinalIgnoreCase))
        {
            _mode = Modes.xlsb;
            OpenXlsb(path, readSharedStrings);
        }
        else
        {
            _mode = Modes.xlsx;
            OpenXlsx(path, updateMode: updateMode);
        }
    }

    private void FillRels(string xmlOrBin)
    {
        var e2 = _xlsxArchive.GetEntry($"xl/_rels/workbook.{xmlOrBin}.rels");
        using var str = e2.Open();
        using var reader = XmlReader.Create(str);
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Whitespace)
            {
                continue;
            }

            if (reader.Name == "Relationship")
            {
                string target = reader.GetAttribute("Target");
                string type = reader.GetAttribute("Type");
                string rId = reader.GetAttribute("Id");

                if (type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
                {
                    _worksheetIdToLocation[rId] = target;
                }
                else if (type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition")
                {
                    _pivotCacheRidToLocation ??= new Dictionary<string, string>();
                    _pivotCacheRidToLocation[rId] = target;
                }
                else if (type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")
                {
                    _sharedStringsLocation = target;
                }
                else if (type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
                {
                    _stylesLocation = target;
                }
                //else if (type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
                //{
                //    _themeLocation = target;
                //}

            }
        }
    }

    /// <summary>
    /// xlsb read strategy, true = More RAM needed but faster
    /// </summary>
    public bool UseMemoryStreamInXlsb = true;
    private static byte[] ReadEntryToBytes(Stream streamToRead, long length)
    {
        byte[] byteArray = new byte[length];
        int bytesRead = 0;
        int toRead = 65_536;
        try
        {
            while (true)
            {
                if (length - bytesRead < toRead)
                {
                    toRead = (int)length - bytesRead;
                }

                int pos = streamToRead.Read(byteArray, bytesRead, toRead);
                bytesRead += pos;
                if (pos == 0)
                {
                    break;
                }
            }
            return byteArray;
        }
        finally
        {
            streamToRead.Dispose();
        }
    }

    public void FillSharedStrings()
    {
        if (_sharedStringsLocation == null)
        {
            throw new Exception("no shared strings found");
        }

        var sharedstringsEntry = _xlsxArchive.GetEntry($@"xl/{_sharedStringsLocation}");
        Stream str = sharedstringsEntry.Open();
        //Stream str = getMemoryStream(sharedstringsEntry.Open(), sharedstringsEntry.Length);

        try
        {
            var xmlSettings = new XmlReaderSettings
            {
                IgnoreComments = true,
                IgnoreWhitespace = true,
                NameTable = new SharedStringsNameTable(),
                //CheckCharacters = false,
                //ValidationType = ValidationType.None,
                //ValidationFlags = System.Xml.Schema.XmlSchemaValidationFlags.None
            };
            using var reader = XmlReader.Create(str, xmlSettings);
            reader.Read();
            if (reader.IsStartElement("sst", _openXmlInfoString))
            {
                string unqCnt = reader.GetAttribute("uniqueCount");
                if (unqCnt != null)
                {
                    _ = int.TryParse(unqCnt, out _uniqueStringCount);
                }

                string cnt = reader.GetAttribute("count");
            }
            else
            {
                throw new Exception("not openXml 2006 format!");
            }

            if (_uniqueStringCount !=-1)
            {
                _sharedStringArray = new string[_uniqueStringCount];
            }
            else
            {
                _sharedStringArray = new string[1024];
            }

            int stringNum = 0;
            while (reader.Read())
            {
                if (reader.IsStartElement("si"))
                {
                    reader.Read();
                    if (reader.IsStartElement("t"))
                    {
                        string vall = reader.ReadElementContentAsString();
                        // si -> t->  t -> value
                        //string vall = reader.Value;

                        if (stringNum >= _sharedStringArray.Length)
                        {
                            Array.Resize(ref _sharedStringArray, 2 * _sharedStringArray.Length);
                        }

                        _sharedStringArray[stringNum++] = vall;
                    }
                }
            }

            if (stringNum != _sharedStringArray.Length)
            {
                Array.Resize(ref _sharedStringArray, stringNum);
            }

        }
        finally
        {
            str.Dispose();
        }

    }

    public void FillBinSharedStrings()
    {
        if (_sharedStringsLocation == null)
        {
            throw new Exception("no shared strings found");
        }

        var sharedstringsEntry = _xlsxArchive.GetEntry($@"xl/{_sharedStringsLocation}");

        BiffReaderWriter reader;
        Stream? str = null;
        if (UseMemoryStreamInXlsb)
        {
            reader = new BiffReaderWriter(ReadEntryToBytes(sharedstringsEntry.Open(), sharedstringsEntry.Length));
        }
        else
        {
            str = new BufferedStream(sharedstringsEntry.Open());
            reader = new BiffReaderWriter(str);
        }

        try
        {
            int stringNum = 0;

            reader.ReadSharedStrings();
            if (_sharedStringArray == null && reader._sharedStringUniqueCount != 0)
            {
                _sharedStringArray = new string[reader._sharedStringUniqueCount];
            }
            else
            {
                _sharedStringArray = new string[1024];
            }

            while (reader.ReadSharedStrings())
            {
                string vall = reader._sharedStringValue;
                if (vall != null)
                {
                    //_sharedStringList.Add(vall);
                    if (stringNum >= _sharedStringArray.Length)
                    {
                        Array.Resize(ref _sharedStringArray, 2 * _sharedStringArray.Length);
                    }
                    _sharedStringArray[stringNum++] = vall;
                }
            }
            if (stringNum != _sharedStringArray.Length)
            {
                Array.Resize(ref _sharedStringArray, stringNum);
            }

        }
        finally
        {
            str?.Dispose();
        }

    }

    private void FillStyles()
    {
        var _stylesCellXfs = new List<StyleInfo>();

        var e = _xlsxArchive.GetEntry($"xl/{_stylesLocation}");
        using (var str = e.Open())
        {
            using var reader = XmlReader.Create(str);
            reader.Read();
            while (!reader.EOF)
            {
                if (reader.IsStartElement("cellXfs"))
                {
                    reader.Read();
                    while (!reader.EOF)
                    {
                        if (reader.IsStartElement("xf"))
                        {
                            _ = int.TryParse(reader.GetAttribute("xfId"), out var xfId);
                            _ = int.TryParse(reader.GetAttribute("numFmtId"), out var numFmtId);
                            _stylesCellXfs.Add(new StyleInfo() { XfId = xfId, NumFmtId = numFmtId/*, ApplyNumberFormat = applyNumberFormat*/ });
                            reader.Skip();
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                else if (reader.IsStartElement("numFmts"))
                {
                    reader.Read();
                    while (!reader.EOF)
                    {
                        if (reader.IsStartElement("numFmt"))
                        {
                            string? formatCode = reader.GetAttribute("formatCode");
                            _ = int.TryParse(reader.GetAttribute("numFmtId"), out var numFmtId);

                            if (!_numberFormatsTypeDictionary.TryGetValue(numFmtId, out var type))
                            {
                                if (_dateExcelMasks.Contains(formatCode))
                                {
                                    type = typeof(DateTime?);
                                }
                                else
                                {
                                    type = typeof(string);
                                }
                                _numberFormatsTypeDictionary[numFmtId] = type;
                            }
                            else
                            {
                                _numberFormatsTypeDictionary[numFmtId] = type;
                            }
                            reader.Skip();
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                if (reader.Depth > 0)
                {
                    reader.Skip();
                }
                else
                {
                    reader.Read();
                }
            }
        }
        _stylesCellXfsArray = _stylesCellXfs.ToArray();
    }

    private void FillBinStyles()
    {
        var _stylesCellXfs = new List<StyleInfo>();

        var e = _xlsxArchive.GetEntry($"xl/{_stylesLocation}");
        using (var str = e.Open())
        {
            bool stylesFirstTime = true;
            bool formatFirstTime = true;

            using var reader = new BiffReaderWriter(str);
            while (reader.ReadStyles())
            {
                if (reader._inCellXf)
                {
                    if (stylesFirstTime)
                    {
                        reader.ReadStyles();
                        stylesFirstTime = false;
                    }
                    int numFmtId = reader._numberFormatIndex;
                    int xfId = reader._parentCellStyleXf;
                    _stylesCellXfs.Add(new StyleInfo() { XfId = xfId, NumFmtId = numFmtId });
                    //reader.ReadStyles();
                }
                else if (reader._inNumberFormat)
                {
                    if (formatFirstTime)
                    {
                        reader.ReadStyles();
                        formatFirstTime = false;
                    }

                    string formatCode = reader._formatString;
                    int numFmtId = reader._format;

                    if (!_numberFormatsTypeDictionary.TryGetValue(numFmtId, out var type))
                    {
                        if (_dateExcelMasks.Contains(formatCode))
                        {
                            type = typeof(DateTime?);
                        }
                        else
                        {
                            type = typeof(string);
                        }
                        _numberFormatsTypeDictionary[numFmtId] = type;
                    }
                    //reader.ReadStyles();
                }
            }
        }
        _stylesCellXfsArray = _stylesCellXfs.ToArray();
    }

    public override string[] GetSheetNames()
    {
        return _worksheetIdToName.Values.ToArray();
    }

    private static readonly char[] _digits = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };

    private string _actualSheetDimensions = null;

    private Stream _sheetStream;
    private XmlReader? _xmlReader;
    private BiffReaderWriter _biffReader;
    private ZipArchiveEntry _sheetEntry;

    private int _prevRowNum = -1;
    private bool _isFirstRow = true;
    private int _columnsCntFromFirstRow = -1;
    private int _numberOfFirstColumnWithData = -1;
    private int _numberOfLastColumnWithData = -1;
    private int _colNum = -1;
    private int _prevColNum = -1;
    private int _howManyEmptyRow = -1;
    private int _emptyRowCnt = -1;
    private bool _success = false;
    private int _len = 0;
    private int _rowNum = 0;
    private bool _returnValue = true;

    private void SheetPreInitialize()
    {
        _prevRowNum = -1;
        //isFirstRow = true;
        _columnsCntFromFirstRow = -1;
        _numberOfFirstColumnWithData = -1;
        _numberOfLastColumnWithData = -1;
        _colNum = -1;
        _prevColNum = -1;
        _howManyEmptyRow = -1;
        _emptyRowCnt = -1;
        _success = false;
        _len = 0;
        _rowNum = 0;
        _returnValue = true;
        innerRow = new FieldInfo[4096];
        _sheetEntry = GetArchiverEntry(ActualSheetName);
    }
    private void InitSheetXlsxReader()
    {
        SheetPreInitialize();
        //sheetStream = sheetEntry.Open();
        _sheetStream = new BufferedStream(_sheetEntry.Open(), 65_536);

        var xmlSettings = new XmlReaderSettings
        {
            IgnoreComments = true,
            IgnoreWhitespace = true,
            NameTable = new SheetNameTable(),
            //CheckCharacters = false,
            //ValidationType = ValidationType.None,
            //ValidationFlags = System.Xml.Schema.XmlSchemaValidationFlags.None
        };

        _xmlReader = XmlReader.Create(_sheetStream, xmlSettings);
        _xmlReader.Read();
        while (!_xmlReader.EOF)
        {
            if (_xmlReader.IsStartElement("sheetData"))
            {
                _prevRowNum = -1;
                return;
            }
            else if (_xmlReader.IsStartElement("dimension"))
            {
                _actualSheetDimensions = _xmlReader.GetAttribute("ref");
                _xmlReader.Skip();
            }
            else if (_xmlReader.Depth == 0)
            {
                _xmlReader.Read();
            }
            else
            {
                _xmlReader.Skip();
            }
        }
    }

    private bool ReadXlsx()
    {
        if (_isFirstRow)
        {
            InitSheetXlsxReader();
        }
        if (_emptyRowCnt < _howManyEmptyRow)
        {
            _emptyRowCnt++;
            _prevRowNum++;
            return true;
        }

        //https://github.com/KrzysztofDusko/SpreadSheetTasks/issues/4
        if (innerRow.Length > _numberOfLastColumnWithData - _numberOfFirstColumnWithData)
        {
            for (int i = _numberOfFirstColumnWithData; i < _numberOfLastColumnWithData; i++)
            {
                innerRow[i - _numberOfFirstColumnWithData].type = ExcelDataType.Null;
            }
        }

        if (_emptyRowCnt == -1)
        {
            if (!_xmlReader!.ReadToFollowing("row"))
            {
                _xmlReader.Dispose();
                _sheetStream.Dispose();
                _isFirstRow = true; // RESET
                return false;
            }

            _success = _xmlReader.MoveToAttribute("r");
            if (_success)
            {
                _len = _xmlReader.ReadValueChunk(_buffer, 0, _buffer.Length);
                _rowNum = ParseToUnsignedIntFromBuffer(_buffer, _len);
            }
            else
            {
                // If "r" attribute is missing, increment from previous row
                _rowNum = _prevRowNum + 1;
            }
        }
        _howManyEmptyRow = -1;
        _emptyRowCnt = -1;

        //empty row/s
        if (_prevRowNum != -1 && _rowNum > _prevRowNum + 1)
        {
            for (int i = 0; i < _columnsCntFromFirstRow; i++)
            {
                innerRow[i].type = ExcelDataType.Null;
            }
            _howManyEmptyRow = _rowNum - _prevRowNum - 1;
            _prevRowNum++;
            _emptyRowCnt = 1;
            return true;
        }

        _prevRowNum = _rowNum;
        
        // Clear innerRow for the new row - set all cells to Null
        for (int i = 0; i < _columnsCntFromFirstRow; i++)
        {
            innerRow[i].type = ExcelDataType.Null;
        }

        _prevColNum = -1;
        int cellIndex = 0; // Track cell position sequentially
        while (_xmlReader.Read() && _xmlReader.IsStartElement("c"))
        {
            // Try to get "r" attribute for column position
            bool hasRAttribute = _xmlReader.MoveToAttribute("r");
            
            if (hasRAttribute)
            {
                _len = _xmlReader.ReadValueChunk(_buffer, 0, _buffer.Length);
                _colNum = -1;

                // Fast path for most common references like "A12", "B3", etc.
                if (_len > 1)
                {
                    int v = _buffer[0] - 'A';
                    if ((uint)v < 26u && (uint)(_buffer[1] - '0') <= 9u)
                    {
                        _colNum = v;
                    }
                }

                if (_colNum == -1)
                {
                    //Sylvan
                    for (int j = 0; j < _len; j++)
                    {
                        char c = _buffer[j];
                        if (c < 'A' || c > 'Z')
                        {
                            break;
                        }
                        int v = c - 'A';
                        if ((uint)v < 26u)
                        {
                            _colNum = ((_colNum + 1) * 26) + v;
                        }
                    }
                }
            }
            else
            {
                // If no "r" attribute, use sequential position
                _colNum = cellIndex;
            }
            cellIndex++;

            if (_isFirstRow && _numberOfFirstColumnWithData == -1)
            {
                _numberOfFirstColumnWithData = _colNum;
            }
            if (_isFirstRow && _numberOfLastColumnWithData < _colNum)
            {
                _numberOfLastColumnWithData = _colNum;
            }

            int _tmpLen = ResizeRowIfNeeded();
            ref FieldInfo valueX = ref innerRow[_tmpLen];

            bool isEmptyElement = _xmlReader.IsEmptyElement;
            if (!isEmptyElement)
            {
                char sstMark = '\0';
                int sstLen = 0;
                int styleId = -1;

                _success = _xmlReader.MoveToNextAttribute();

                if (_success && _xmlReader.Name == "s")
                {
                    _len = _xmlReader.ReadValueChunk(_buffer, 0, _buffer.Length);
                    styleId = ParseToUnsignedIntFromBuffer(_buffer, _len);
                    _success = _xmlReader.MoveToNextAttribute();
                }

                if (_success && _xmlReader.Name == "t")
                {
                    sstLen = _xmlReader.ReadValueChunk(_buffer, 0, _buffer.Length);
                    sstMark = _buffer[0];
                }

                _xmlReader.MoveToElement();
                //bool success = false;
                if (sstMark == 'i' && sstLen == 9)
                {
                    //success =
                    _xmlReader.ReadToDescendant("is");
                    _xmlReader.Read();
                    _xmlReader.Read();
                    valueX.type = ExcelDataType.String;
                    valueX.strValue = _xmlReader.ReadContentAsString();
                }
                else if (_xmlReader.ReadToDescendant("v"))
                {
                    _xmlReader.Read();

                    if (sstMark == 's' && sstLen == 1) // 's' = string/sharedstring'b' = boolean, 'e' = error
                    {
                        _len = _xmlReader.ReadValueChunk(_buffer, 0, _buffer.Length);
                        valueX.type = ExcelDataType.String;
                        valueX.strValue = _sharedStringArray[ParseToUnsignedIntFromBuffer(_buffer, _len)];
                    }
                    else if (sstMark == 's' || sstMark == 'i' && sstLen == 9)  // InlineStr?
                    {
                        if (_xmlReader.NodeType != XmlNodeType.EndElement || _xmlReader.Name != "c")
                        {
                            //if (xmlReader.NodeType != XmlNodeType.Text)
                            //{ 
                            //    xmlReader.Read();
                            //}
                            valueX.type = ExcelDataType.String;
                            valueX.strValue = _xmlReader.ReadContentAsString();
                        }
                        else
                        {
                            valueX.type = ExcelDataType.Null;
                        }
                        //valueX = _buffer.AsSpan(0, len).ToString();
                    }
                    else if (TreatAllColumnsAsText)
                    {
                        valueX.type = ExcelDataType.String;
                        valueX.strValue = _xmlReader.ReadContentAsString();
                    }
                    else
                    {
                        _len = _xmlReader.ReadValueChunk(_buffer, 0, _buffer.Length);
                        if (sstMark == 'b') // 'b' = boolean, 'e' = error
                        {
                            valueX.type = ExcelDataType.Boolean;
                            //valueX.int32Value = (_buffer[0] - '0');
                            valueX.boolValue = _buffer[0] == '1';
                        }
                        else if (sstMark == 'e') // 'b' = boolean, 'e' = error
                        {
                            valueX.type = ExcelDataType.String;
                            valueX.strValue = "error in cell";
                        }
                        else if (styleId != -1)
                        {
                            var s = _stylesCellXfsArray[styleId];
                            int numFormatId = s.NumFmtId;

                            if (_numberFormatsTypeDictionary.TryGetValue(numFormatId, out Type type) && type == typeof(DateTime?)
                                && double.TryParse(_buffer.AsSpan(0, _len), NumberStyles.Any, provider: invariantCultureInfo, out double doubleDate)
                                //&& FastDoubleParser.TryParseDouble(_buffer.AsSpan(0, len), out double doubleDate)
                                )
                            {
                                valueX.type = ExcelDataType.DateTime;
                                valueX.dtValue = DateTime.FromOADate(doubleDate);
                            }
                            else
                            {
                                if (ContainsDoubleMarks(_buffer, _len))
                                {
                                    valueX.type = ExcelDataType.Double;
                                    valueX.doubleValue = double.Parse(_buffer.AsSpan(0, _len), provider: invariantCultureInfo);
                                    //valueX.doubleValue = FastDoubleParser.ParseDouble(_buffer.AsSpan(0, len));
                                }
                                else
                                {
                                    valueX.type = ExcelDataType.Int64;
                                    valueX.int64Value = ParseToInt64FromBuffer(_buffer, _len);
                                }
                            }
                        }
                        else
                        {
                            if (ContainsDoubleMarks(_buffer, _len))
                            {
                                valueX.type = ExcelDataType.Double;
                                valueX.doubleValue = double.Parse(_buffer.AsSpan(0, _len), provider: invariantCultureInfo);
                                //valueX.doubleValue = FastDoubleParser.ParseDouble(_buffer.AsSpan(0, len));
                            }
                            else
                            {
                                valueX.type = ExcelDataType.Int64;
                                valueX.int64Value = ParseToInt64FromBuffer(_buffer, _len);
                            }
                        }
                    }
                }
                else
                {
                    valueX.type = ExcelDataType.Null;
                }
            }

            _prevColNum = _colNum;

            if (!isEmptyElement) // depth = ...
            {
                while (_xmlReader.Depth > 3)
                {
                    //reader.Skip();
                    _xmlReader.Read();
                }
            }
        }

        if (_isFirstRow)
        {
            _isFirstRow = false;
            _columnsCntFromFirstRow = _numberOfLastColumnWithData - _numberOfFirstColumnWithData + 1;
            Array.Resize<FieldInfo>(ref innerRow, _columnsCntFromFirstRow);
            FieldCount = _columnsCntFromFirstRow;
        }
        return true;
    }

    private int ResizeRowIfNeeded()
    {
        int _tmpLen = _colNum - _numberOfFirstColumnWithData;
        if (_tmpLen >= innerRow.Length)
        {
            Array.Resize(ref innerRow, _tmpLen + 1);
        }
        return _tmpLen;
    }

    private void InitSheetXlsbReader()
    {
        SheetPreInitialize();
        if (_sheetEntry.Length >= int.MaxValue)
        {
            UseMemoryStreamInXlsb = false;
        }
        if (UseMemoryStreamInXlsb)
        {
            _biffReader = new BiffReaderWriter(ReadEntryToBytes(_sheetEntry.Open(), _sheetEntry.Length));
        }
        else
        {
            _sheetStream = new BufferedStream(_sheetEntry.Open(), 65_536);
            _biffReader = new BiffReaderWriter(_sheetStream);
        }

        while (!_biffReader._readCell) //read to cell
        {
            _biffReader.ReadWorksheet();
        }
        //prevColNum = colNum;
        _rowNum = _biffReader._rowIndex;
        _colNum = _biffReader._columnNum;
    }

    private bool ReadXlsb()
    {
        //fist time = initialize
        if (_isFirstRow)
        {
            InitSheetXlsbReader();
        }

        //last row is not complete
        if (!_isFirstRow && _colNum > _numberOfFirstColumnWithData)
        {
            for (int i = _numberOfFirstColumnWithData; i < _colNum; i++)
            {
                innerRow[i - _numberOfFirstColumnWithData].type = ExcelDataType.Null;
            }
        }
        //previous read = false
        if (!_returnValue)
        {
            _biffReader.Dispose();
            _sheetStream?.Dispose();
            _isFirstRow = true; // RESET
            return false;
        }

        //missing rows
        if (_prevRowNum != -1 && _rowNum > _prevRowNum + 1)
        {
            for (int i = 0; i < _columnsCntFromFirstRow; i++)
            {
                innerRow[i].type = ExcelDataType.Null;
            }
            _howManyEmptyRow = _rowNum - _prevRowNum - 1;
            _prevRowNum++;
            _emptyRowCnt = 1;
            return true;
        }
        _prevRowNum = _rowNum;

        // Clear innerRow for the new row - set all cells to Null
        for (int i = 0; i < _columnsCntFromFirstRow; i++)
        {
            innerRow[i].type = ExcelDataType.Null;
        }

        while (_rowNum == _prevRowNum && _returnValue)
        {
            if (_biffReader._readCell)
            {
                //determine first row length
                if (_isFirstRow)
                {
                    if (_numberOfFirstColumnWithData == -1)
                    {
                        _numberOfFirstColumnWithData = _colNum;
                    }
                    _numberOfLastColumnWithData = _colNum;
                    _columnsCntFromFirstRow = _numberOfLastColumnWithData - _numberOfFirstColumnWithData + 1;
                    FieldCount = _columnsCntFromFirstRow;
                }

                int _tmpLen = ResizeRowIfNeeded();
                ref FieldInfo valueX = ref innerRow[_tmpLen];

                valueX.type = ExcelDataType.String;
                switch (_biffReader._cellType)
                {
                    case CellType.sharedString:
                        valueX.type = ExcelDataType.String;
                        valueX.strValue = _sharedStringArray[_biffReader._intValue];
                        break;
                    case CellType.stringVal:
                        valueX.type = ExcelDataType.String;
                        valueX.strValue = _biffReader._stringValue;
                        break;
                    case CellType.boolVal:
                        if (!TreatAllColumnsAsText)
                        {
                            valueX.type = ExcelDataType.Boolean;
                            //valueX.int32Value = biffReader.boolValue ? 1 : 0;
                            valueX.boolValue = _biffReader._boolValue;
                        }
                        else
                        {
                            valueX.type = ExcelDataType.String;
                            valueX.strValue = _biffReader._boolValue.ToString();
                        }
                        break;
                    case CellType.doubleVal:
                        {
                            double doubleVal = _biffReader._doubleVal;
                            var styleIndex = _biffReader._xfIndex;
                            if (styleIndex == 0) // general
                            {
                                SetValueForXlsb(doubleVal, ref valueX);
                                if (TreatAllColumnsAsText)
                                {
                                    valueX.type = ExcelDataType.String;
                                    valueX.strValue = valueX.doubleValue.ToString();
                                }
                            }
                            else
                            {
                                int numFormatId = _stylesCellXfsArray[styleIndex].NumFmtId;
                                if (_numberFormatsTypeDictionary.TryGetValue(numFormatId, out Type type) && type == typeof(DateTime?))
                                {
                                    valueX.type = ExcelDataType.DateTime;
                                    valueX.dtValue = DateTime.FromOADate((double)doubleVal);
                                    if (TreatAllColumnsAsText)
                                    {
                                        valueX.type = ExcelDataType.String;
                                        valueX.strValue = valueX.dtValue.ToString();
                                    }
                                }
                                else
                                {
                                    XlsxOrXlsbReadOrEdit.SetValueForXlsb(doubleVal, ref valueX);
                                    if (TreatAllColumnsAsText)
                                    {
                                        valueX.type = ExcelDataType.String;
                                        valueX.strValue = valueX.doubleValue.ToString();
                                    }
                                }
                            }
                            break;
                        }
                    default:
                        valueX.type = ExcelDataType.Null;
                        break;
                }
            }

            _biffReader.ReadWorksheet();
            while (_returnValue && !_biffReader._readCell)
            {
                _returnValue = _biffReader.ReadWorksheet();
            }

            _prevColNum = _colNum;
            _rowNum = _biffReader._rowIndex;
            _colNum = _biffReader._columnNum;

            if (!_isFirstRow)
                {
                    if (_colNum > _prevColNum + 1 && _rowNum == _prevRowNum)
                    {
                        //prevColNum  = prev not null
                        //colNum = current
                        // A B !NULL! !NULL! C
                        for (int i = _prevColNum + 1; i < _colNum; i++)
                        {
                            innerRow[i - _numberOfFirstColumnWithData].type = ExcelDataType.Null;
                        }
                    }
                    // A B C !NULL! !NULL!
                    else if (_rowNum > _prevRowNum && _prevColNum < _numberOfLastColumnWithData)
                    {
                        for (int i = 1; i <= _numberOfLastColumnWithData - _prevColNum; i++)
                        {
                            innerRow[_prevColNum + i - _numberOfFirstColumnWithData].type = ExcelDataType.Null;
                        }
                    }
                }
        }

        if (_isFirstRow)
        {
            _isFirstRow = false;
            Array.Resize(ref innerRow, _columnsCntFromFirstRow);
        }

        return true;
    }

    public override bool Read()
    {
        if (_mode == Modes.xlsx)
        {
            return ReadXlsx();
        }
        else
        {
            return ReadXlsb();
        }
    }

    //inspired by https://github.com/MarkPflug/Sylvan.Data.Excel
    private readonly static char[] _buffer = new char[64];

    private static int ParseToUnsignedIntFromBuffer(char[] buff, int len)
    {
        int res = 0;
        for (int i = 0; i < len; i++)
        {
            res = res * 10 + (buff[i] - '0');
        }
        return res;
    }

    private static Int64 ParseToInt64FromBuffer(char[] buff, int len)
    {
        //Int64 res = 0;
        //int start = buff[0] == '-' ? 1 : 0;
        //for (int i = start; i < len; i++)
        //{
        //    res = res * 10 + (buff[i] - '0');
        //}
        //return start == 1 ? -res : res;
        int start = buff[0] == '-' ? 1 : 0;
        if (len <= start)
        {
            return 0; // Invalid input - just a minus sign or empty
        }
        Int64 res = buff[start] - '0';
        for (int i = start + 1; i < len; i++)
        {
            res = res * 10 + (buff[i] - '0');
        }
        return start == 1 ? -res : res;
    }

    private static bool ContainsDoubleMarks(char[] buff, int len)
    {
        //for (int i = 0; i < len; i++)
        //{
        //    char c = buff[i];
        //    if (c == '.' || c == 'E')
        //    {
        //        return true;
        //    }
        //}
        //return false;
        return buff.AsSpan(0,len).IndexOfAny('.', 'E') > 0;
    }

    public IEnumerable<object[]> GetRowsOfXlsx(string sheetName)
    {
        ActualSheetName = sheetName;
        Read();
        object[] row = new object[FieldCount];
        GetValues(row);
        yield return row;

        while (ReadXlsx())
        {
            GetValues(row);
            yield return row;
        }
    }

    public IEnumerable<object[]> GetRowsOfXlsb(string sheetName)
    {
        ActualSheetName = sheetName;
        Read();
        object[] row = new object[FieldCount];
        GetValues(row);
        yield return row;

        while (ReadXlsb())
        {
            GetValues(row);
            yield return row;
        }
    }

    public IEnumerable<object[]> GetRowsOfSheet(string sheetName)
    {
        ArgumentNullException.ThrowIfNull(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);

        if (_mode == Modes.xlsx)
        {
            return GetRowsOfXlsx(sheetName);
        }
        else
        {
            return GetRowsOfXlsb(sheetName);
        }
    }

    private ZipArchiveEntry GetArchiverEntry(string sheetName)
    {
        string id = _worksheetNameToId[sheetName];
        string location = _worksheetIdToLocation[id];
        return _xlsxArchive.GetEntry($"xl/{location}");
    }

    private static (int row, int column) GetNumbersFromAdress(string startingCellAdress)
    {
        int n1 = startingCellAdress.IndexOfAny(_digits);
        string letters = startingCellAdress[..n1];
        int rowNumFromAdress = int.Parse(startingCellAdress[n1..]);
        int colNumFromAdress = _letterToColumnNum[letters] + 1;
        return (rowNumFromAdress - 1, colNumFromAdress - 1);
    }

    /// <summary>
    /// clear and fill existing tab with dataReader
    /// </summary>
    /// <param name="sheetName"></param>
    /// <param name="reader"></param>
    /// <param name="startingCellAdress"></param>
    /// <returns></returns>
    public string ReplaceSheetData(string sheetName, IDataReader reader, string startingCellAdress = "A1")
    {
        ArgumentNullException.ThrowIfNull(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentNullException.ThrowIfNull(reader);
        ArgumentException.ThrowIfNullOrWhiteSpace(startingCellAdress);

        if (!_worksheetNameToId.TryGetValue(sheetName, out string id))
        {
            throw new Exception($"ReplaceSheetData - {sheetName} not found");
        }

        string location = _worksheetIdToLocation[id];
        var sheetEntryToReplace = _xlsxArchive.GetEntry($"xl/{location}");
        sheetEntryToReplace.Delete();

        (int startingRow, int startingColumn) = GetNumbersFromAdress(startingCellAdress);

        int _dateStyleNum = 0;

        for (int i = 0; i < _stylesCellXfsArray.Length; i++)
        {
            var item = _stylesCellXfsArray[i];
            int numFormatId = item.NumFmtId;
            if (_numberFormatsTypeDictionary.TryGetValue(numFormatId, out Type type) && type == typeof(DateTime?))
            {
                _dateStyleNum = i;
                break;
            }
        }

        string excelRangeString = "";
        var newData = _xlsxArchive.CreateEntry($"xl/{location}");
        using (StreamWriter sheetData = new FormattingStreamWriter(newData.Open(), Encoding.UTF8, CultureInfo.InvariantCulture.NumberFormat))
        {
            int cnt = XlsxWriter.WriteToExisting(sheetData, reader, dateStyleNum: _dateStyleNum, dateTimeStyleNum: _dateStyleNum, startingRow, startingColumn);
            string start = XlsxWriter._letters[startingColumn] + (startingRow + 1).ToString();
            string end = XlsxWriter._letters[startingColumn + reader.FieldCount - 1] + (startingRow + 1 + cnt).ToString();
            excelRangeString = start + ":" + end;
        }
        return excelRangeString;
    }

    private List<string> GetPivotTableList()
    {
        List<string> pivotTableList = new List<string>();
        var e = _xlsxArchive.GetEntry("[Content_Types].xml");
        using (var str = e.Open())
        {
            using var reader = XmlReader.Create(str);
            while (reader.Read())
            {
                if (reader.Name == "Override")
                {
                    string nm = reader.GetAttribute("ContentType");
                    if (nm == "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml")
                    {
                        pivotTableList.Add(reader.GetAttribute("PartName"));
                    }
                }
            }
        }
        return pivotTableList;
    }
    private int GetCacheIdForPivotTable(string pivotTableName, List<string> pivotTableList)
    {
        int cacheId = -1;
        foreach (var pivotTableLocation in pivotTableList)
        {
            var e1 = _xlsxArchive.GetEntry(pivotTableLocation[1..]);
            using var str = e1.Open();
            using var reader = XmlReader.Create(str);
            while (reader.Read())
            {
                if (reader.IsStartElement("pivotTableDefinition"))
                {
                    string name = reader.GetAttribute("name");
                    if (name == pivotTableName)
                    {
                        _ = int.TryParse(reader.GetAttribute("cacheId"), out cacheId);
                        break;
                    }
                }
            }
        }

        return cacheId;
    }

    /// <summary>
    /// changes data reference for pivot table, use with ReplaceSheetData
    /// </summary>
    /// <param name="pivotTableName">name of pivot table</param>
    /// <param name="referention">new reference</param>
    /// <param name="doRefreshOnLoad">refresh pivot table on open</param>
    public void ReplacePivotTableDim(string pivotTableName, string referention, bool doRefreshOnLoad = true)
    {
        ArgumentNullException.ThrowIfNull(pivotTableName);
        ArgumentException.ThrowIfNullOrWhiteSpace(pivotTableName);
        ArgumentNullException.ThrowIfNull(referention);
        ArgumentException.ThrowIfNullOrWhiteSpace(referention);

        var pivotTableList = GetPivotTableList();
        int cacheId = GetCacheIdForPivotTable(pivotTableName, pivotTableList);

        if (cacheId == -1)
        {
            throw new Exception($"{pivotTableName} - not found");
        }

        string rId = _pivotCacheIdToRid[cacheId];
        string tableCacheLocation = _pivotCacheRidToLocation[rId];
        var pivotCacheEntry = _xlsxArchive.GetEntry("xl/" + tableCacheLocation);

        bool addRefreshOnLoad = false;
        bool replaceRefresh = false;
        string reff = null;
        string pivotTableXmlAsPlainTxt = null;
        using (var str = pivotCacheEntry.Open())
        {
            using var reader = XmlReader.Create(str);
            while (reader.Read())
            {
                if (reader.IsStartElement("pivotCacheDefinition"))
                {
                    string refreshOnLoad = reader.GetAttribute("refreshOnLoad");
                    if (refreshOnLoad == null)
                    {
                        addRefreshOnLoad = true;
                    }
                    else if (refreshOnLoad == "0")
                    {
                        replaceRefresh = true;
                    }
                }
                else if (reader.IsStartElement("cacheSource"))
                {
                    do
                    {
                        reader.Read();
                    } while (!reader.IsStartElement("worksheetSource"));
                    reff = reader.GetAttribute("ref");
                    break;
                }
            }
        }

        if (reff != null)
        {
            using var str = pivotCacheEntry.Open();
            using var reader = new StreamReader(str);
            pivotTableXmlAsPlainTxt = reader.ReadToEnd();
        }
        
        if (pivotTableXmlAsPlainTxt == null)
        {
            throw new Exception("Pivot table cache source not found");
        }
        
        pivotCacheEntry.Delete();

        int firsPartIndex = pivotTableXmlAsPlainTxt.IndexOf(@"</cacheSource>");
        string firsPartTxt = pivotTableXmlAsPlainTxt[0..firsPartIndex];

        var ent = _xlsxArchive.CreateEntry("xl/" + tableCacheLocation);
        using (var str = ent.Open())
        {
            using var sw = new StreamWriter(str);
            if (doRefreshOnLoad && addRefreshOnLoad)
            {
                firsPartTxt = firsPartTxt.Replace(" r:id=", " refreshOnLoad=\"1\" r:id=");
            }
            else if (doRefreshOnLoad && replaceRefresh)
            {
                firsPartTxt = firsPartTxt.Replace("refreshOnLoad=\"0\" r:id=", "refreshOnLoad=\"1\" r:id=");
            }

            firsPartTxt = firsPartTxt.Replace($"ref=\"{reff}\"", $"ref=\"{referention}\"");

            sw.Write(firsPartTxt);
            sw.Write(pivotTableXmlAsPlainTxt.AsSpan()[firsPartIndex..]);
        }
    }

    private string _actualSheetName;
    public override string ActualSheetName
    {
        get => _actualSheetName;
        set
        {
            // Reset reader state when switching sheets
            if (_actualSheetName != value)
            {
                _isFirstRow = true;
                _prevRowNum = -1;
                _columnsCntFromFirstRow = -1;
                _numberOfFirstColumnWithData = -1;
                _numberOfLastColumnWithData = -1;
                _colNum = -1;
                _prevColNum = -1;
                _howManyEmptyRow = -1;
                _emptyRowCnt = -1;
                _rowNum = 0;
                _returnValue = true;
                innerRow = new FieldInfo[4096];
            }
            _rowCount = -2;
            _actualSheetName = value;
        }
    }

    private int _resultCount = -1;
    public override int ResultsCount { get => _resultCount; }
    //private string Name { get => ActualSheetName; }

    private int _rowCount = -2;

    /// <summary>
    /// row count (not always available)
    /// </summary>
    public override int RowCount { get => _rowCount != -2 ? _rowCount : PrepareRowCnt(); }

    private int PrepareRowCnt()
    {
        _rowCount = -1;
        if (_actualSheetDimensions == null)
        {
            _rowCount = -1;
            return _rowCount;
        }

        int i1 = _actualSheetDimensions.IndexOf(':');
        string t1 = _actualSheetDimensions[..i1];
        int i2 = t1.IndexOfAny(_digits);
        _ = int.TryParse(t1[i2..], out int start);

        t1 = _actualSheetDimensions[(i1 + 1)..];
        i2 = t1.IndexOfAny(_digits);

        _ = int.TryParse(t1[i2..], out int end);
        _rowCount = end - start; // header is not row !!

        return _rowCount;
    }

    private static void SetValueForXlsb(double rawValue, ref FieldInfo fieldInfo)
    {
        long l1 = (long)rawValue;
        double res = l1 - /*(double)*/rawValue;
        if (res < 3 * double.Epsilon && res > -3 * double.Epsilon)
        {
            fieldInfo.type = ExcelDataType.Int64;
            fieldInfo.int64Value = l1;
        }
        else
        {
            fieldInfo.type = ExcelDataType.Double;
            fieldInfo.doubleValue = rawValue;
        }
    }
}

//https://github.com/MarkPflug/Sylvan.Data.Excel/blob/main/source/Sylvan.Data.Excel/Xlsx/SheetNameTable.cs#L13
sealed class SheetNameTable : NameTable
{
    public override string Add(char[] key, int start, int len)
    {
        return Get(key.AsSpan(start, len)) ?? base.Add(key, start, len);
    }

    public override string Add(string key)
    {
        return Get(key.AsSpan()) ?? base.Add(key);
    }

    public override string? Get(char[] key, int start, int len)
    {
        return Get(key.AsSpan(start, len));
    }

    public override string? Get(string value)
    {
        return Get(value.AsSpan());
    }

    public string? Get(ReadOnlySpan<char> value)
    {
        switch (value.Length)
        {
            case 0:
                return string.Empty;
            case 1:
                switch (value[0])
                {
                    case 'c': return "c";
                    case 'r': return "r";
                    case 't': return "t";
                    case 's': return "s";
                    case 'v': return "v";
                }
                break;
            case 2:
                if (value.SequenceEqual("is")) return "is";
                break;
            case 3:
                if (value.SequenceEqual("row")) return "row";
                if (value.SequenceEqual("ref")) return "ref";
                if (value.SequenceEqual("col")) return "col";
                if (value.SequenceEqual("min")) return "min";
                if (value.SequenceEqual("max")) return "max";
                break;
            case 4:
                if (value.SequenceEqual("cols")) return "cols";
                break;
            case 5:
                if (value.SequenceEqual("spans")) return "spans";
                break;
            case 6:
                if (value.SequenceEqual("hidden")) return "hidden";
                break;
            case 9:
                if (value.SequenceEqual("dyDescent")) return "dyDescent";
                if (value.SequenceEqual("dimension")) return "dimension";
                if (value.SequenceEqual("sheetData")) return "sheetData";
                break;
        }
        return null;
    }
}

//https://github.com/MarkPflug/Sylvan.Data.Excel/blob/206bd17d06cc06edafbecd7fa8868a9b700b6e7a/source/Sylvan.Data.Excel/Xlsx/SheetNameTable.cs#L80
sealed class SharedStringsNameTable : NameTable
{
    public override string Add(char[] key, int start, int len)
    {
        return Get(key.AsSpan(start, len)) ?? base.Add(key, start, len);
    }

    public override string Add(string key)
    {
        return Get(key.AsSpan()) ?? base.Add(key);
    }

    public override string? Get(char[] key, int start, int len)
    {
        return Get(key.AsSpan(start, len));
    }

    public override string? Get(string value)
    {
        return Get(value.AsSpan());
    }

    public string? Get(ReadOnlySpan<char> value)
    {
        switch (value.Length)
        {
            case 0:
                return string.Empty;
            case 1:
                if (value.SequenceEqual("t")) return "t";
                break;
            case 2:
                if (value.SequenceEqual("si")) return "si";
                break;
            case 3:
                if (value.SequenceEqual("sst")) return "sst";
                break;
            case 5:
                if (value.SequenceEqual("count")) return "count";
                break;
            case 11:
                if (value.SequenceEqual("uniqueCount")) return "uniqueCount";
                break;
        }
        return null;
    }
}
