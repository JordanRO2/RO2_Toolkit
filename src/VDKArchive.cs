using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace VDKTool
{
    public class FileEntry
    {
        public string Name { get; set; }
        public string Path { get; set; }
        public bool IsDirectory { get; set; }
        public uint UncompressedSize { get; set; }
        public uint CompressedSize { get; set; }
        public uint Offset { get; set; }
        public long DataPosition { get; set; }
    }

    public class VDKArchive
    {
        private const int ENTRY_SIZE = 145;
        private const int NAME_SIZE = 128;
        private static readonly Encoding KoreanEncoding;

        public string FilePath { get; private set; }
        public string Version { get; private set; }
        public uint FileCount { get; private set; }
        public uint FolderCount { get; private set; }
        public List<FileEntry> Entries { get; private set; }

        static VDKArchive()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            KoreanEncoding = Encoding.GetEncoding(51949); // euc-kr
        }

        public VDKArchive()
        {
            Entries = new List<FileEntry>();
        }

        public static VDKArchive Load(string filePath)
        {
            var archive = new VDKArchive { FilePath = filePath };

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            using (var reader = new BinaryReader(stream))
            {
                // Read header (24 bytes)
                byte[] versionBytes = reader.ReadBytes(8);
                int nullPos = Array.IndexOf(versionBytes, (byte)0);
                if (nullPos < 0) nullPos = 8;
                archive.Version = Encoding.ASCII.GetString(versionBytes, 0, nullPos);

                uint magic = reader.ReadUInt32();
                archive.FileCount = reader.ReadUInt32();
                archive.FolderCount = reader.ReadUInt32();
                uint totalSize = reader.ReadUInt32();

                // Validate
                if (archive.Version == "VDISK1.0")
                {
                    if (magic != 4294967040)
                        throw new InvalidDataException("Invalid VDISK1.0 magic");
                }
                else if (archive.Version == "VDISK1.1")
                {
                    uint validation = reader.ReadUInt32();
                    uint expected = archive.FileCount * 264 + 4;
                    if (validation != expected)
                        throw new InvalidDataException("Invalid VDISK1.1 validation");
                }
                else
                {
                    throw new InvalidDataException($"Unknown VDK format: {archive.Version}");
                }

                // Parse entries recursively
                ParseEntries(reader, "", archive.Entries);
            }

            return archive;
        }

        private static void ParseEntries(BinaryReader reader, string currentPath, List<FileEntry> entries)
        {
            while (true)
            {
                if (reader.BaseStream.Position + ENTRY_SIZE > reader.BaseStream.Length)
                    break;

                byte[] entryData = reader.ReadBytes(ENTRY_SIZE);

                bool isDir = entryData[0] != 0;

                // Extract name (bytes 1-128)
                byte[] nameBytes = new byte[NAME_SIZE];
                Array.Copy(entryData, 1, nameBytes, 0, NAME_SIZE);
                int nameEnd = Array.IndexOf(nameBytes, (byte)0);
                if (nameEnd < 0) nameEnd = NAME_SIZE;
                string name = KoreanEncoding.GetString(nameBytes, 0, nameEnd);

                uint uncompSize = BitConverter.ToUInt32(entryData, 129);
                uint compSize = BitConverter.ToUInt32(entryData, 133);
                uint offset = BitConverter.ToUInt32(entryData, 141);

                string fullPath = string.IsNullOrEmpty(currentPath) ? name : Path.Combine(currentPath, name);
                long dataPosition = reader.BaseStream.Position;

                var entry = new FileEntry
                {
                    Name = name,
                    Path = fullPath,
                    IsDirectory = isDir,
                    UncompressedSize = uncompSize,
                    CompressedSize = compSize,
                    Offset = offset,
                    DataPosition = dataPosition
                };
                entries.Add(entry);

                if (isDir)
                {
                    if (name != "." && name != "..")
                        ParseEntries(reader, fullPath, entries);
                }
                else
                {
                    reader.BaseStream.Seek(dataPosition + compSize, SeekOrigin.Begin);
                }

                if (offset == 0)
                    break;
            }
        }

        public byte[] ExtractFile(FileEntry entry)
        {
            using (var stream = new FileStream(FilePath, FileMode.Open, FileAccess.Read))
            {
                stream.Seek(entry.DataPosition, SeekOrigin.Begin);
                byte[] compressedData = new byte[entry.CompressedSize];
                stream.Read(compressedData, 0, (int)entry.CompressedSize);

                if (entry.UncompressedSize == entry.CompressedSize)
                    return compressedData;

                // Try to decompress
                try
                {
                    return DecompressZlib(compressedData);
                }
                catch
                {
                    // Try raw deflate
                    try
                    {
                        return DecompressDeflate(compressedData);
                    }
                    catch
                    {
                        return compressedData;
                    }
                }
            }
        }

        private static byte[] DecompressZlib(byte[] data)
        {
            // Skip zlib header (2 bytes)
            using (var input = new MemoryStream(data, 2, data.Length - 2))
            using (var deflate = new DeflateStream(input, CompressionMode.Decompress))
            using (var output = new MemoryStream())
            {
                deflate.CopyTo(output);
                return output.ToArray();
            }
        }

        private static byte[] DecompressDeflate(byte[] data)
        {
            using (var input = new MemoryStream(data))
            using (var deflate = new DeflateStream(input, CompressionMode.Decompress))
            using (var output = new MemoryStream())
            {
                deflate.CopyTo(output);
                return output.ToArray();
            }
        }

        public List<FileEntry> GetFileEntries()
        {
            return Entries.FindAll(e => !e.IsDirectory && e.Name != "." && e.Name != "..");
        }

        public List<FileEntry> GetDirectoryEntries()
        {
            return Entries.FindAll(e => e.IsDirectory && e.Name != "." && e.Name != "..");
        }
    }

    public class VDKWriter
    {
        private const int ENTRY_SIZE = 145;
        private const int NAME_SIZE = 128;
        private const int HEADER_SIZE = 28;
        private static readonly Encoding KoreanEncoding;

        private Dictionary<string, byte[]> files = new Dictionary<string, byte[]>();
        private HashSet<string> dirs = new HashSet<string>();
        private List<(string path, long offset)> fileEntries = new List<(string, long)>();
        private Dictionary<string, byte[]> compressedCache = new Dictionary<string, byte[]>();
        private bool compress = true;

        static VDKWriter()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            KoreanEncoding = Encoding.GetEncoding(51949);
        }

        public void AddFile(string archivePath, byte[] data)
        {
            archivePath = archivePath.Replace('\\', '/');
            files[archivePath] = data;

            // Track directories
            string[] parts = archivePath.Split('/');
            for (int i = 0; i < parts.Length - 1; i++)
            {
                string dirPath = string.Join("/", parts, 0, i + 1);
                dirs.Add(dirPath);
            }
        }

        public void AddDirectory(string sourceDir, Action<int, int, string> progressCallback = null)
        {
            var allFiles = Directory.GetFiles(sourceDir, "*", SearchOption.AllDirectories);
            int count = 0;

            foreach (var filePath in allFiles)
            {
                string relativePath = filePath.Substring(sourceDir.Length).TrimStart('\\', '/');
                byte[] data = File.ReadAllBytes(filePath);
                AddFile(relativePath, data);

                count++;
                progressCallback?.Invoke(count, allFiles.Length, relativePath);
            }
        }

        private byte[] GetCompressedData(string path, byte[] data)
        {
            if (compressedCache.TryGetValue(path, out byte[] cached))
                return cached;

            if (compress && data.Length > 0)
            {
                byte[] compressed = CompressZlib(data);
                if (compressed.Length < data.Length)
                {
                    compressedCache[path] = compressed;
                    return compressed;
                }
            }

            compressedCache[path] = data;
            return data;
        }

        private long CalcNodeSize(Dictionary<string, object> node, string pathPrefix)
        {
            long size = 0;

            // . entry
            size += ENTRY_SIZE;
            // .. entry
            size += ENTRY_SIZE;

            var dirDict = (Dictionary<string, object>)node["__dirs__"];
            var fileList = (List<(string, byte[])>)node["__files__"];

            var sortedDirs = new List<string>(dirDict.Keys);
            sortedDirs.Sort(StringComparer.OrdinalIgnoreCase);

            // Subdirectories
            foreach (var subdir in sortedDirs)
            {
                size += ENTRY_SIZE;
                string subpath = string.IsNullOrEmpty(pathPrefix) ? subdir : $"{pathPrefix}/{subdir}";
                size += CalcNodeSize((Dictionary<string, object>)dirDict[subdir], subpath);
            }

            // Files
            foreach (var (fileName, fileData) in fileList)
            {
                string filepath = string.IsNullOrEmpty(pathPrefix) ? fileName : $"{pathPrefix}/{fileName}";
                byte[] compData = GetCompressedData(filepath, fileData);
                size += ENTRY_SIZE + compData.Length;
            }

            return size;
        }

        public int Write(string outputPath, bool compress = true)
        {
            this.compress = compress;
            fileEntries.Clear();
            compressedCache.Clear();

            var root = BuildTree();
            int filesCount = files.Count;
            int foldersCount = dirs.Count;  // Just named directories

            using (var stream = new FileStream(outputPath, FileMode.Create, FileAccess.ReadWrite))
            using (var writer = new BinaryWriter(stream))
            {
                // Write placeholder header (28 bytes)
                writer.Write(new byte[HEADER_SIZE]);

                // Write root . entry
                var rootDirs = new List<string>(((Dictionary<string, object>)root["__dirs__"]).Keys);
                rootDirs.Sort(StringComparer.OrdinalIgnoreCase);

                long rootDotOffset = stream.Position;
                if (rootDirs.Count > 0)
                {
                    long nextPos = rootDotOffset + ENTRY_SIZE;
                    WriteDirEntry(writer, ".", (uint)nextPos);

                    // Write root level directories
                    for (int i = 0; i < rootDirs.Count; i++)
                    {
                        bool isLast = (i == rootDirs.Count - 1);
                        var dirDict = (Dictionary<string, object>)root["__dirs__"];
                        WriteDirectoryRecursive(writer, rootDirs[i],
                            (Dictionary<string, object>)dirDict[rootDirs[i]], isLast, rootDirs[i]);
                    }
                }
                else
                {
                    WriteDirEntry(writer, ".", 0);
                }

                // Calculate hierarchical data size
                long hierDataEnd = stream.Position;
                uint hierSize = (uint)(hierDataEnd - HEADER_SIZE);

                // Write flat table at the end
                long flatTableStart = stream.Position;
                writer.Write((uint)fileEntries.Count);

                foreach (var (filePath, entryOffset) in fileEntries)
                {
                    // Uppercase path, 260 bytes padded with nulls
                    string upperPath = filePath.ToUpperInvariant().Replace('\\', '/');
                    byte[] pathBytes = KoreanEncoding.GetBytes(upperPath);
                    byte[] paddedPath = new byte[260];
                    int copyLen = Math.Min(pathBytes.Length, 259);
                    Array.Copy(pathBytes, 0, paddedPath, 0, copyLen);
                    writer.Write(paddedPath);
                    writer.Write((uint)entryOffset);
                }

                uint flatTableSize = (uint)(stream.Position - flatTableStart);

                // Go back and write the real header
                stream.Seek(0, SeekOrigin.Begin);
                writer.Write(Encoding.ASCII.GetBytes("VDISK1.1"));
                writer.Write((uint)0);  // Magic (unused)
                writer.Write((uint)filesCount);
                writer.Write((uint)foldersCount);
                writer.Write(hierSize);
                writer.Write(flatTableSize);
            }

            return filesCount;
        }

        private Dictionary<string, object> BuildTree()
        {
            var root = new Dictionary<string, object>
            {
                ["__files__"] = new List<(string, byte[])>(),
                ["__dirs__"] = new Dictionary<string, object>()
            };

            foreach (var kvp in files)
            {
                string[] parts = kvp.Key.Split('/');
                var current = root;

                for (int i = 0; i < parts.Length - 1; i++)
                {
                    var dirsDict = (Dictionary<string, object>)current["__dirs__"];
                    if (!dirsDict.ContainsKey(parts[i]))
                    {
                        dirsDict[parts[i]] = new Dictionary<string, object>
                        {
                            ["__files__"] = new List<(string, byte[])>(),
                            ["__dirs__"] = new Dictionary<string, object>()
                        };
                    }
                    current = (Dictionary<string, object>)dirsDict[parts[i]];
                }

                ((List<(string, byte[])>)current["__files__"]).Add((parts[parts.Length - 1], kvp.Value));
            }

            return root;
        }

        private void WriteDirEntry(BinaryWriter writer, string name, uint nextOffset)
        {
            byte[] entry = new byte[ENTRY_SIZE];
            entry[0] = 1; // is_dir

            byte[] nameBytes = KoreanEncoding.GetBytes(name);
            int copyLen = Math.Min(nameBytes.Length, NAME_SIZE - 1);
            Array.Copy(nameBytes, 0, entry, 1, copyLen);

            // sizes = 0 for dirs
            Array.Copy(BitConverter.GetBytes(nextOffset), 0, entry, 141, 4);

            writer.Write(entry);
        }

        private void WriteFileEntry(BinaryWriter writer, string name, byte[] data, byte[] compData,
                                    uint nextOffset, string fullPath)
        {
            long entryOffset = writer.BaseStream.Position;

            // Record for flat table
            fileEntries.Add((fullPath, entryOffset));

            byte[] entry = new byte[ENTRY_SIZE];
            entry[0] = 0; // not dir

            byte[] nameBytes = KoreanEncoding.GetBytes(name);
            int copyLen = Math.Min(nameBytes.Length, NAME_SIZE - 1);
            Array.Copy(nameBytes, 0, entry, 1, copyLen);

            Array.Copy(BitConverter.GetBytes((uint)data.Length), 0, entry, 129, 4);
            Array.Copy(BitConverter.GetBytes((uint)compData.Length), 0, entry, 133, 4);
            // data_offset at 137 = 0 (data follows immediately)
            Array.Copy(BitConverter.GetBytes(nextOffset), 0, entry, 141, 4);

            writer.Write(entry);
            writer.Write(compData);
        }

        private void WriteDirectoryRecursive(BinaryWriter writer, string name,
                                             Dictionary<string, object> node, bool isLastSibling, string pathPrefix)
        {
            // Calculate content size for sibling offset
            long contentSize = CalcNodeSize(node, pathPrefix);
            long currentPos = writer.BaseStream.Position;
            uint dirOffset = isLastSibling ? 0 : (uint)(currentPos + ENTRY_SIZE + contentSize);

            // Write directory name entry
            WriteDirEntry(writer, name, dirOffset);

            // Write . entry - points to next entry (..)
            long dotPos = writer.BaseStream.Position;
            WriteDirEntry(writer, ".", (uint)(dotPos + ENTRY_SIZE));

            // Get children
            var dirDict = (Dictionary<string, object>)node["__dirs__"];
            var fileList = (List<(string, byte[])>)node["__files__"];

            var sortedDirs = new List<string>(dirDict.Keys);
            sortedDirs.Sort(StringComparer.OrdinalIgnoreCase);
            fileList.Sort((a, b) => string.Compare(a.Item1, b.Item1, StringComparison.OrdinalIgnoreCase));

            bool hasChildren = sortedDirs.Count > 0 || fileList.Count > 0;

            // Write .. entry - points to first child or 0 if empty
            if (hasChildren)
            {
                long dotdotPos = writer.BaseStream.Position;
                WriteDirEntry(writer, "..", (uint)(dotdotPos + ENTRY_SIZE));
            }
            else
            {
                WriteDirEntry(writer, "..", 0);
            }

            // Write subdirectories
            for (int i = 0; i < sortedDirs.Count; i++)
            {
                bool isLast = (i == sortedDirs.Count - 1) && fileList.Count == 0;
                string subpath = $"{pathPrefix}/{sortedDirs[i]}";
                WriteDirectoryRecursive(writer, sortedDirs[i],
                    (Dictionary<string, object>)dirDict[sortedDirs[i]], isLast, subpath);
            }

            // Write files
            for (int i = 0; i < fileList.Count; i++)
            {
                var (fileName, fileData) = fileList[i];
                string filepath = $"{pathPrefix}/{fileName}";
                byte[] compData = GetCompressedData(filepath, fileData);

                uint nextOffset;
                if (i == fileList.Count - 1)
                {
                    nextOffset = 0;
                }
                else
                {
                    nextOffset = (uint)(writer.BaseStream.Position + ENTRY_SIZE + compData.Length);
                }

                WriteFileEntry(writer, fileName, fileData, compData, nextOffset, filepath);
            }
        }

        private static byte[] CompressZlib(byte[] data)
        {
            using (var output = new MemoryStream())
            {
                // Write zlib header
                output.WriteByte(0x78);
                output.WriteByte(0x9C);

                using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, true))
                {
                    deflate.Write(data, 0, data.Length);
                }

                // Calculate Adler32 checksum
                uint adler = Adler32(data);
                output.WriteByte((byte)((adler >> 24) & 0xFF));
                output.WriteByte((byte)((adler >> 16) & 0xFF));
                output.WriteByte((byte)((adler >> 8) & 0xFF));
                output.WriteByte((byte)(adler & 0xFF));

                return output.ToArray();
            }
        }

        private static uint Adler32(byte[] data)
        {
            uint a = 1, b = 0;
            foreach (byte d in data)
            {
                a = (a + d) % 65521;
                b = (b + a) % 65521;
            }
            return (b << 16) | a;
        }
    }
}
