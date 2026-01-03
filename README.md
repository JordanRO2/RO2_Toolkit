# RO2 Toolkit

C# toolkit for Ragnarok Online 2 game data management.

## Features

- **VDK Archive Management**: Extract and pack VDK archives (VDISK format)
- **CT Table Conversion**: Convert CT binary tables to/from Excel (XLSX/CSV)
- **Windows GUI + CLI**

## Usage

### GUI
```
bin/VDK_Tool.exe
```

### CLI
```bash
# VDK Operations
VDK_Tool.exe extract <file.vdk> [output_dir]
VDK_Tool.exe extractall <dir> [suffix]
VDK_Tool.exe list <file.vdk>

# CT Conversion
VDK_Tool.exe ct2xlsx <file.ct>
VDK_Tool.exe ct2csv <file.ct>
VDK_Tool.exe ctall <dir>
```

## Project Structure

```
RO2-Toolkit/
├── bin/
│   ├── VDK_Tool.exe
│   └── *.dll
├── src/
│   ├── Program.cs
│   ├── MainForm.cs
│   ├── VDKArchive.cs
│   └── CTProcessor.cs
└── build.bat
```

## Supported Formats

### VDK Archive (VDISK 1.0/1.1)
- Hierarchical file storage
- Zlib/Deflate compression
- EUC-KR filename encoding

### CT Binary Table
- RO2 game data format
- 10 data types (BYTE, SHORT, INT, DWORD, STRING, FLOAT, etc.)
- CRC-16 XMODEM checksum
- UTF-16LE string encoding

## Building

Requires .NET Framework 4.8:
```bash
build.bat
```

## Dependencies

- ClosedXML 0.102.2
- System.Text.Encoding.CodePages 7.0.0
