using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ClosedXML.Excel;

namespace VDKTool
{
    /// <summary>
    /// CT (Custom Table) file processor for Ragnarok Online 2.
    /// Supports reading and writing CT binary format.
    /// </summary>
    public class CTProcessor
    {
        private const int CT_HEADER_SIZE = 64;
        private const string CT_MAGIC_NEW = "RO2SEC!";
        private const string CT_MAGIC_OLD = "RO2!";
        private string _detectedMagic = CT_MAGIC_NEW;

        // CT data types
        public enum CTDataType
        {
            BYTE = 2,
            SHORT = 3,
            WORD = 4,
            INT = 5,
            DWORD = 6,
            DWORD_HEX = 7,
            STRING = 8,
            FLOAT = 9,
            INT64 = 11,
            BOOL = 12
        }

        public static readonly Dictionary<int, string> TypeNames = new Dictionary<int, string>
        {
            { 2, "BYTE" },
            { 3, "SHORT" },
            { 4, "WORD" },
            { 5, "INT" },
            { 6, "DWORD" },
            { 7, "DWORD_HEX" },
            { 8, "STRING" },
            { 9, "FLOAT" },
            { 11, "INT64" },
            { 12, "BOOL" }
        };

        public static readonly Dictionary<string, int> TypeCodes = new Dictionary<string, int>
        {
            { "BYTE", 2 },
            { "SHORT", 3 },
            { "WORD", 4 },
            { "INT", 5 },
            { "DWORD", 6 },
            { "DWORD_HEX", 7 },
            { "STRING", 8 },
            { "FLOAT", 9 },
            { "INT64", 11 },
            { "BOOL", 12 }
        };

        public string FilePath { get; private set; }
        public List<string> Headers { get; private set; }
        public List<string> Types { get; private set; }
        public List<List<string>> Rows { get; private set; }
        public string Timestamp { get; private set; }

        public CTProcessor()
        {
            Headers = new List<string>();
            Types = new List<string>();
            Rows = new List<List<string>>();
        }

        /// <summary>
        /// Read a CT file and parse its contents.
        /// </summary>
        public void Read(string filePath)
        {
            FilePath = filePath;
            Headers.Clear();
            Types.Clear();
            Rows.Clear();

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            using (var reader = new BinaryReader(stream))
            {
                // Read and validate header
                ReadHeader(reader);

                // Read columns
                int numColumns = reader.ReadInt32();
                for (int i = 0; i < numColumns; i++)
                {
                    Headers.Add(ReadString(reader));
                }

                // Read types
                int numTypes = reader.ReadInt32();
                for (int i = 0; i < numTypes; i++)
                {
                    int typeCode = reader.ReadInt32();
                    Types.Add(TypeNames.ContainsKey(typeCode) ? TypeNames[typeCode] : $"UNKNOWN_{typeCode}");
                }

                // Read rows
                int numRows = reader.ReadInt32();
                for (int i = 0; i < numRows; i++)
                {
                    var row = new List<string>();
                    for (int j = 0; j < Types.Count; j++)
                    {
                        row.Add(ReadValue(reader, Types[j]));
                    }
                    Rows.Add(row);
                }
            }
        }

        /// <summary>
        /// Write data to a CT file.
        /// </summary>
        public void Write(string filePath, List<string> headers, List<string> types, List<List<string>> rows, string timestamp = null)
        {
            using (var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            using (var writer = new BinaryWriter(stream))
            {
                // Write header
                WriteHeader(writer, timestamp ?? DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                // Write columns
                writer.Write(headers.Count);
                foreach (var header in headers)
                {
                    WriteString(writer, header);
                }

                // Write types
                writer.Write(types.Count);
                foreach (var typeName in types)
                {
                    writer.Write(TypeCodes.ContainsKey(typeName) ? TypeCodes[typeName] : 5);
                }

                // Capture row data for CRC
                using (var rowDataStream = new MemoryStream())
                using (var rowDataWriter = new BinaryWriter(rowDataStream))
                {
                    // Write row count to main stream
                    writer.Write(rows.Count);

                    // Write rows
                    for (int i = 0; i < rows.Count; i++)
                    {
                        var row = rows[i];
                        for (int j = 0; j < types.Count; j++)
                        {
                            string value = j < row.Count ? row[j] : "";
                            byte[] valueBytes = PackValue(types[j], value);
                            writer.Write(valueBytes);
                            rowDataWriter.Write(valueBytes);
                        }
                    }

                    // Calculate and write CRC
                    byte[] rowData = rowDataStream.ToArray();
                    ushort crc = CalculateCRC16(rowData);
                    writer.Write(crc);
                }
            }
        }

        private void ReadHeader(BinaryReader reader)
        {
            byte[] headerData = reader.ReadBytes(CT_HEADER_SIZE);

            // Try new magic first, then old magic
            byte[] newMagicBytes = Encoding.Unicode.GetBytes(CT_MAGIC_NEW);
            byte[] oldMagicBytes = Encoding.Unicode.GetBytes(CT_MAGIC_OLD);
            byte[] magicBytes = null;

            bool isNewMagic = true;
            for (int i = 0; i < newMagicBytes.Length && i < headerData.Length; i++)
            {
                if (headerData[i] != newMagicBytes[i])
                {
                    isNewMagic = false;
                    break;
                }
            }

            if (isNewMagic)
            {
                magicBytes = newMagicBytes;
                _detectedMagic = CT_MAGIC_NEW;
            }
            else
            {
                // Try old magic
                bool isOldMagic = true;
                for (int i = 0; i < oldMagicBytes.Length && i < headerData.Length; i++)
                {
                    if (headerData[i] != oldMagicBytes[i])
                    {
                        isOldMagic = false;
                        break;
                    }
                }

                if (isOldMagic)
                {
                    magicBytes = oldMagicBytes;
                    _detectedMagic = CT_MAGIC_OLD;
                }
                else
                {
                    throw new InvalidDataException("Invalid CT file magic (expected RO2SEC! or RO2!)");
                }
            }

            // Extract timestamp
            int timestampStart = magicBytes.Length + 2;
            int timestampEnd = timestampStart;
            while (timestampEnd < headerData.Length - 1)
            {
                if (headerData[timestampEnd] == 0 && headerData[timestampEnd + 1] == 0)
                    break;
                timestampEnd += 2;
            }

            if (timestampEnd > timestampStart)
            {
                byte[] timestampBytes = new byte[timestampEnd - timestampStart];
                Array.Copy(headerData, timestampStart, timestampBytes, 0, timestampBytes.Length);
                Timestamp = Encoding.Unicode.GetString(timestampBytes);
            }
        }

        private void WriteHeader(BinaryWriter writer, string timestamp)
        {
            byte[] header = new byte[CT_HEADER_SIZE];

            // Write magic (use detected magic from read, or default to new)
            byte[] magicBytes = Encoding.Unicode.GetBytes(_detectedMagic);
            Array.Copy(magicBytes, header, magicBytes.Length);

            // Add null terminator after magic
            int pos = magicBytes.Length;
            header[pos++] = 0;
            header[pos++] = 0;

            // Write timestamp
            byte[] timestampBytes = Encoding.Unicode.GetBytes(timestamp);
            Array.Copy(timestampBytes, 0, header, pos, Math.Min(timestampBytes.Length, CT_HEADER_SIZE - pos - 2));

            writer.Write(header);
        }

        private string ReadString(BinaryReader reader)
        {
            int length = reader.ReadInt32();
            if (length == 0) return "";

            byte[] bytes = reader.ReadBytes(length * 2);
            return Encoding.Unicode.GetString(bytes);
        }

        private void WriteString(BinaryWriter writer, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                writer.Write(0);
                return;
            }

            writer.Write(value.Length);
            writer.Write(Encoding.Unicode.GetBytes(value));
        }

        private string ReadValue(BinaryReader reader, string typeName)
        {
            switch (typeName)
            {
                case "BYTE":
                case "BOOL":
                    return reader.ReadByte().ToString();
                case "SHORT":
                    return reader.ReadInt16().ToString();
                case "WORD":
                    return reader.ReadUInt16().ToString();
                case "INT":
                    return reader.ReadInt32().ToString();
                case "DWORD":
                    return reader.ReadUInt32().ToString();
                case "DWORD_HEX":
                    return "0x" + reader.ReadUInt32().ToString("X");
                case "FLOAT":
                    return reader.ReadSingle().ToString();
                case "INT64":
                    return reader.ReadInt64().ToString();
                case "STRING":
                    return ReadString(reader);
                default:
                    return reader.ReadInt32().ToString();
            }
        }

        private byte[] PackValue(string typeName, string value)
        {
            using (var ms = new MemoryStream())
            using (var writer = new BinaryWriter(ms))
            {
                if (string.IsNullOrEmpty(value)) value = "0";

                switch (typeName)
                {
                    case "BYTE":
                    case "BOOL":
                        writer.Write(byte.Parse(value));
                        break;
                    case "SHORT":
                        writer.Write(short.Parse(value));
                        break;
                    case "WORD":
                        writer.Write(ushort.Parse(value));
                        break;
                    case "INT":
                        writer.Write(int.Parse(value));
                        break;
                    case "DWORD":
                        writer.Write(uint.Parse(value));
                        break;
                    case "DWORD_HEX":
                        if (value.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
                            writer.Write(Convert.ToUInt32(value, 16));
                        else
                            writer.Write(uint.Parse(value));
                        break;
                    case "FLOAT":
                        writer.Write(float.Parse(value));
                        break;
                    case "INT64":
                        writer.Write(long.Parse(value));
                        break;
                    case "STRING":
                        if (string.IsNullOrEmpty(value) || value == "0")
                        {
                            writer.Write(0);
                        }
                        else
                        {
                            writer.Write(value.Length);
                            writer.Write(Encoding.Unicode.GetBytes(value));
                        }
                        break;
                    default:
                        writer.Write(int.Parse(value));
                        break;
                }

                return ms.ToArray();
            }
        }

        private ushort CalculateCRC16(byte[] data)
        {
            ushort crc = 0x0000;
            ushort poly = 0x1021;

            foreach (byte b in data)
            {
                crc ^= (ushort)(b << 8);
                for (int i = 0; i < 8; i++)
                {
                    if ((crc & 0x8000) != 0)
                        crc = (ushort)((crc << 1) ^ poly);
                    else
                        crc = (ushort)(crc << 1);
                }
            }

            return crc;
        }

        /// <summary>
        /// Export CT data to XLSX format (matching Python RO2-Table-Converter format).
        /// Row 1: Types (light blue), Row 2: Headers (dark blue), Row 3+: Data
        /// </summary>
        public void ExportToXLSX(string outputPath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("RO2 Table Data");

                // Row 1: Types (light blue background)
                for (int i = 0; i < Types.Count; i++)
                {
                    var cell = worksheet.Cell(1, i + 1);
                    cell.Value = Types[i];
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#D9E2F3");
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }

                // Row 2: Headers (dark blue background, white text)
                for (int i = 0; i < Headers.Count; i++)
                {
                    var cell = worksheet.Cell(2, i + 1);
                    cell.Value = Headers[i];
                    cell.Style.Font.Bold = true;
                    cell.Style.Font.FontColor = XLColor.White;
                    cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#366092");
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }

                // Row 3+: Data rows
                for (int rowIdx = 0; rowIdx < Rows.Count; rowIdx++)
                {
                    var row = Rows[rowIdx];
                    for (int colIdx = 0; colIdx < row.Count; colIdx++)
                    {
                        var cell = worksheet.Cell(rowIdx + 3, colIdx + 1);
                        string value = row[colIdx];

                        // Convert value based on type
                        if (colIdx < Types.Count)
                        {
                            switch (Types[colIdx])
                            {
                                case "FLOAT":
                                    if (float.TryParse(value, out float f))
                                        cell.Value = f;
                                    else
                                        cell.Value = value;
                                    break;
                                case "INT":
                                case "SHORT":
                                case "BYTE":
                                case "BOOL":
                                    if (int.TryParse(value, out int intVal))
                                        cell.Value = intVal;
                                    else
                                        cell.Value = value;
                                    break;
                                case "DWORD":
                                case "WORD":
                                case "INT64":
                                    if (long.TryParse(value, out long l))
                                        cell.Value = l;
                                    else
                                        cell.Value = value;
                                    break;
                                case "DWORD_HEX":
                                    // Keep hex values as strings
                                    cell.Value = value;
                                    break;
                                default:
                                    cell.Value = value;
                                    break;
                            }
                        }
                        else
                        {
                            cell.Value = value;
                        }
                    }
                }

                // Freeze first two rows (types + headers)
                worksheet.SheetView.FreezeRows(2);

                // Auto-fit columns (min 10, max 50)
                foreach (var col in worksheet.ColumnsUsed())
                {
                    col.AdjustToContents();
                    if (col.Width < 10) col.Width = 10;
                    if (col.Width > 50) col.Width = 50;
                }

                workbook.SaveAs(outputPath);
            }
        }

        /// <summary>
        /// Export CT data to CSV format.
        /// </summary>
        public void ExportToCSV(string outputPath)
        {
            using (var writer = new StreamWriter(outputPath, false, Encoding.UTF8))
            {
                // Write headers
                writer.WriteLine(string.Join(",", Headers.ConvertAll(EscapeCSV)));

                // Write types as comment
                writer.WriteLine("#" + string.Join(",", Types));

                // Write rows
                foreach (var row in Rows)
                {
                    writer.WriteLine(string.Join(",", row.ConvertAll(EscapeCSV)));
                }
            }
        }

        /// <summary>
        /// Import CT data from XLSX format.
        /// Row 1: Types, Row 2: Headers, Row 3+: Data
        /// Always writes as new format (RO2SEC!) when saving.
        /// </summary>
        public void ImportFromXLSX(string inputPath)
        {
            Headers.Clear();
            Types.Clear();
            Rows.Clear();
            _detectedMagic = CT_MAGIC_NEW; // Always use new format for output

            using (var workbook = new XLWorkbook(inputPath))
            {
                var worksheet = workbook.Worksheet(1);
                var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;
                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;

                if (lastColumn == 0 || lastRow < 2)
                    throw new InvalidDataException("XLSX file must have at least 2 rows (types and headers)");

                // Row 1: Types
                for (int col = 1; col <= lastColumn; col++)
                {
                    string type = worksheet.Cell(1, col).GetString().Trim().ToUpper();
                    if (string.IsNullOrEmpty(type)) type = "INT";
                    Types.Add(type);
                }

                // Row 2: Headers
                for (int col = 1; col <= lastColumn; col++)
                {
                    Headers.Add(worksheet.Cell(2, col).GetString());
                }

                // Row 3+: Data
                for (int row = 3; row <= lastRow; row++)
                {
                    var rowData = new List<string>();
                    for (int col = 1; col <= lastColumn; col++)
                    {
                        var cell = worksheet.Cell(row, col);
                        string value;

                        // Handle different cell types appropriately
                        if (cell.IsEmpty())
                        {
                            value = Types[col - 1] == "STRING" ? "" : "0";
                        }
                        else if (col <= Types.Count && Types[col - 1] == "FLOAT")
                        {
                            // Preserve float precision
                            if (cell.TryGetValue(out double d))
                                value = d.ToString("G");
                            else
                                value = cell.GetString();
                        }
                        else if (col <= Types.Count && Types[col - 1] == "DWORD_HEX")
                        {
                            // Keep hex format if present, otherwise convert
                            value = cell.GetString();
                            if (!value.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
                            {
                                if (uint.TryParse(value, out uint uval))
                                    value = "0x" + uval.ToString("X");
                            }
                        }
                        else
                        {
                            value = cell.GetString();
                        }

                        rowData.Add(value);
                    }
                    Rows.Add(rowData);
                }
            }
        }

        /// <summary>
        /// Import CT data from CSV format.
        /// </summary>
        public void ImportFromCSV(string inputPath)
        {
            Headers.Clear();
            Types.Clear();
            Rows.Clear();

            using (var reader = new StreamReader(inputPath, Encoding.UTF8))
            {
                // Read headers
                string headerLine = reader.ReadLine();
                if (headerLine != null)
                    Headers.AddRange(ParseCSVLine(headerLine));

                // Read types
                string typeLine = reader.ReadLine();
                if (typeLine != null && typeLine.StartsWith("#"))
                    Types.AddRange(ParseCSVLine(typeLine.Substring(1)));

                // Read rows
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (!string.IsNullOrWhiteSpace(line))
                        Rows.Add(ParseCSVLine(line));
                }
            }
        }

        private string EscapeCSV(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            return value;
        }

        private List<string> ParseCSVLine(string line)
        {
            var result = new List<string>();
            bool inQuotes = false;
            var current = new StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"')
                        {
                            current.Append('"');
                            i++;
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        current.Append(c);
                    }
                }
                else
                {
                    if (c == '"')
                    {
                        inQuotes = true;
                    }
                    else if (c == ',')
                    {
                        result.Add(current.ToString());
                        current.Clear();
                    }
                    else
                    {
                        current.Append(c);
                    }
                }
            }

            result.Add(current.ToString());
            return result;
        }
    }
}
