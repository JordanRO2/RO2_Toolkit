using System;
using System.IO;
using System.Windows.Forms;

namespace VDKTool
{
    static class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            // CLI mode if arguments provided
            if (args.Length > 0)
            {
                return RunCLI(args);
            }

            // GUI mode
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
            return 0;
        }

        static int RunCLI(string[] args)
        {
            if (args.Length < 1)
            {
                PrintUsage();
                return 1;
            }

            string command = args[0].ToLower();

            switch (command)
            {
                case "extract":
                case "x":
                    if (args.Length < 2)
                    {
                        Console.WriteLine("Error: VDK file path required");
                        PrintUsage();
                        return 1;
                    }
                    return ExtractVDK(args);

                case "extractall":
                case "xa":
                    if (args.Length < 2)
                    {
                        Console.WriteLine("Error: Directory path required");
                        PrintUsage();
                        return 1;
                    }
                    return ExtractAllVDKs(args);

                case "list":
                case "l":
                    if (args.Length < 2)
                    {
                        Console.WriteLine("Error: VDK file path required");
                        return 1;
                    }
                    return ListVDK(args[1]);

                case "ct2xlsx":
                case "ct":
                    if (args.Length < 2)
                    {
                        Console.WriteLine("Error: CT file or directory path required");
                        PrintUsage();
                        return 1;
                    }
                    return ConvertCT(args, "xlsx");

                case "ct2csv":
                    if (args.Length < 2)
                    {
                        Console.WriteLine("Error: CT file or directory path required");
                        PrintUsage();
                        return 1;
                    }
                    return ConvertCT(args, "csv");

                case "ctall":
                case "cta":
                    if (args.Length < 2)
                    {
                        Console.WriteLine("Error: Directory path required");
                        PrintUsage();
                        return 1;
                    }
                    return ConvertAllCT(args);

                case "xlsx2ct":
                    if (args.Length < 2)
                    {
                        Console.WriteLine("Error: XLSX file path required");
                        PrintUsage();
                        return 1;
                    }
                    return ConvertXLSXtoCT(args[1]);

                case "pack":
                case "p":
                    if (args.Length < 2)
                    {
                        Console.WriteLine("Error: Source directory required");
                        PrintUsage();
                        return 1;
                    }
                    return PackDirectory(args);

                case "help":
                case "-h":
                case "--help":
                    PrintUsage();
                    return 0;

                default:
                    // Assume it's a VDK file path for extraction
                    return ExtractVDK(new[] { "extract", args[0], args.Length > 1 ? args[1] : null });
            }
        }

        static void PrintUsage()
        {
            Console.WriteLine("VDK Tool - Command Line Interface");
            Console.WriteLine();
            Console.WriteLine("VDK Commands:");
            Console.WriteLine("  VDK_Tool.exe                         - Launch GUI");
            Console.WriteLine("  VDK_Tool.exe extract <file.vdk> [output_dir]");
            Console.WriteLine("  VDK_Tool.exe x <file.vdk> [output_dir]");
            Console.WriteLine("                                       - Extract VDK to directory");
            Console.WriteLine("  VDK_Tool.exe extractall <dir> [suffix]");
            Console.WriteLine("  VDK_Tool.exe xa <dir> [suffix]");
            Console.WriteLine("                                       - Extract all VDKs in directory");
            Console.WriteLine("                                         suffix defaults to _UNPACKED");
            Console.WriteLine("  VDK_Tool.exe list <file.vdk>         - List VDK contents");
            Console.WriteLine("  VDK_Tool.exe pack <dir> [output.vdk]  - Pack directory into VDK");
            Console.WriteLine("  VDK_Tool.exe p <dir> [output.vdk]     - Same as pack");
            Console.WriteLine();
            Console.WriteLine("CT Conversion Commands:");
            Console.WriteLine("  VDK_Tool.exe ct2xlsx <file.ct>       - Convert CT to XLSX");
            Console.WriteLine("  VDK_Tool.exe ct <file.ct>            - Same as ct2xlsx");
            Console.WriteLine("  VDK_Tool.exe ct2csv <file.ct>        - Convert CT to CSV");
            Console.WriteLine("  VDK_Tool.exe xlsx2ct <file.xlsx>     - Convert XLSX back to CT");
            Console.WriteLine("  VDK_Tool.exe ctall <dir>             - Convert all CT files in dir to XLSX");
            Console.WriteLine("  VDK_Tool.exe cta <dir>               - Same as ctall");
            Console.WriteLine();
            Console.WriteLine("  VDK_Tool.exe help                    - Show this help");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  VDK_Tool.exe x ASSET.VDK ASSET_UNPACKED");
            Console.WriteLine("  VDK_Tool.exe xa \"C:\\Data\" _UNPACKED");
            Console.WriteLine("  VDK_Tool.exe ct ItemInfo.ct");
            Console.WriteLine("  VDK_Tool.exe ctall \"C:\\Data\\ASSET_UNPACKED\\ASSET\"");
        }

        static int ExtractVDK(string[] args)
        {
            string vdkPath = args[1];

            if (!File.Exists(vdkPath))
            {
                Console.WriteLine($"Error: File not found: {vdkPath}");
                return 1;
            }

            // Determine output directory
            string outputDir;
            if (args.Length >= 3 && !string.IsNullOrEmpty(args[2]))
            {
                outputDir = args[2];
            }
            else
            {
                // Default: same name as VDK with _UNPACKED suffix
                string baseName = Path.GetFileNameWithoutExtension(vdkPath);
                string parentDir = Path.GetDirectoryName(vdkPath);
                outputDir = Path.Combine(parentDir, baseName + "_UNPACKED");
            }

            Console.WriteLine($"Extracting: {vdkPath}");
            Console.WriteLine($"Output: {outputDir}");

            try
            {
                var archive = VDKArchive.Load(vdkPath);
                var files = archive.GetFileEntries();

                Console.WriteLine($"Version: {archive.Version}");
                Console.WriteLine($"Files: {files.Count}");
                Console.WriteLine();

                int extracted = 0;
                int failed = 0;

                foreach (var entry in files)
                {
                    try
                    {
                        string outPath = Path.Combine(outputDir, entry.Path);
                        string outDir = Path.GetDirectoryName(outPath);

                        if (!Directory.Exists(outDir))
                            Directory.CreateDirectory(outDir);

                        byte[] data = archive.ExtractFile(entry);
                        File.WriteAllBytes(outPath, data);

                        extracted++;

                        // Progress every 100 files
                        if (extracted % 100 == 0)
                            Console.WriteLine($"  Extracted {extracted}/{files.Count} files...");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"  Failed: {entry.Path} - {ex.Message}");
                        failed++;
                    }
                }

                Console.WriteLine();
                Console.WriteLine($"Done! Extracted {extracted} files" + (failed > 0 ? $", {failed} failed" : ""));
                return failed > 0 ? 2 : 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return 1;
            }
        }

        static int ExtractAllVDKs(string[] args)
        {
            string directory = args[1];
            string suffix = args.Length >= 3 ? args[2] : "_UNPACKED";

            if (!Directory.Exists(directory))
            {
                Console.WriteLine($"Error: Directory not found: {directory}");
                return 1;
            }

            var vdkFiles = Directory.GetFiles(directory, "*.VDK");
            if (vdkFiles.Length == 0)
            {
                Console.WriteLine($"No VDK files found in: {directory}");
                return 1;
            }

            Console.WriteLine($"Found {vdkFiles.Length} VDK files in: {directory}");
            Console.WriteLine();

            int success = 0;
            int failed = 0;

            foreach (var vdkPath in vdkFiles)
            {
                string baseName = Path.GetFileNameWithoutExtension(vdkPath);
                string outputDir = Path.Combine(directory, baseName + suffix);

                Console.WriteLine($"[{success + failed + 1}/{vdkFiles.Length}] {baseName}.VDK -> {baseName}{suffix}/");

                try
                {
                    var archive = VDKArchive.Load(vdkPath);
                    var files = archive.GetFileEntries();

                    foreach (var entry in files)
                    {
                        string outPath = Path.Combine(outputDir, entry.Path);
                        string outDir = Path.GetDirectoryName(outPath);

                        if (!Directory.Exists(outDir))
                            Directory.CreateDirectory(outDir);

                        byte[] data = archive.ExtractFile(entry);
                        File.WriteAllBytes(outPath, data);
                    }

                    Console.WriteLine($"    Extracted {files.Count} files");
                    success++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"    ERROR: {ex.Message}");
                    failed++;
                }
            }

            Console.WriteLine();
            Console.WriteLine($"Complete! {success} succeeded, {failed} failed");
            return failed > 0 ? 2 : 0;
        }

        static int ListVDK(string vdkPath)
        {
            if (!File.Exists(vdkPath))
            {
                Console.WriteLine($"Error: File not found: {vdkPath}");
                return 1;
            }

            try
            {
                var archive = VDKArchive.Load(vdkPath);
                var files = archive.GetFileEntries();
                var dirs = archive.GetDirectoryEntries();

                Console.WriteLine($"File: {vdkPath}");
                Console.WriteLine($"Version: {archive.Version}");
                Console.WriteLine($"Directories: {dirs.Count}");
                Console.WriteLine($"Files: {files.Count}");
                Console.WriteLine();

                foreach (var entry in files)
                {
                    Console.WriteLine($"  {entry.Path} ({entry.UncompressedSize:N0} bytes)");
                }

                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return 1;
            }
        }

        static int ConvertCT(string[] args, string format)
        {
            string ctPath = args[1];

            if (!File.Exists(ctPath))
            {
                Console.WriteLine($"Error: File not found: {ctPath}");
                return 1;
            }

            string ext = format == "xlsx" ? ".xlsx" : ".csv";
            string outputPath = Path.ChangeExtension(ctPath, ext);

            Console.WriteLine($"Converting: {ctPath}");
            Console.WriteLine($"Output: {outputPath}");

            try
            {
                var processor = new CTProcessor();
                processor.Read(ctPath);

                Console.WriteLine($"Columns: {processor.Headers.Count}");
                Console.WriteLine($"Rows: {processor.Rows.Count}");

                if (format == "xlsx")
                    processor.ExportToXLSX(outputPath);
                else
                    processor.ExportToCSV(outputPath);

                Console.WriteLine($"Done!");
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return 1;
            }
        }

        static int ConvertAllCT(string[] args)
        {
            string directory = args[1];

            if (!Directory.Exists(directory))
            {
                Console.WriteLine($"Error: Directory not found: {directory}");
                return 1;
            }

            // Find all CT files recursively
            var ctFiles = Directory.GetFiles(directory, "*.ct", SearchOption.AllDirectories);
            if (ctFiles.Length == 0)
            {
                Console.WriteLine($"No CT files found in: {directory}");
                return 1;
            }

            Console.WriteLine($"Found {ctFiles.Length} CT files in: {directory}");
            Console.WriteLine();

            int success = 0;
            int failed = 0;

            foreach (var ctPath in ctFiles)
            {
                string relativePath = ctPath.Substring(directory.Length).TrimStart(Path.DirectorySeparatorChar);
                string outputPath = Path.ChangeExtension(ctPath, ".xlsx");

                Console.WriteLine($"[{success + failed + 1}/{ctFiles.Length}] {relativePath}");

                try
                {
                    var processor = new CTProcessor();
                    processor.Read(ctPath);
                    processor.ExportToXLSX(outputPath);

                    Console.WriteLine($"    -> {processor.Headers.Count} columns, {processor.Rows.Count} rows");
                    success++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"    ERROR: {ex.Message}");
                    failed++;
                }
            }

            Console.WriteLine();
            Console.WriteLine($"Complete! {success} succeeded, {failed} failed");
            return failed > 0 ? 2 : 0;
        }

        static int ConvertXLSXtoCT(string xlsxPath)
        {
            if (!File.Exists(xlsxPath))
            {
                Console.WriteLine($"Error: File not found: {xlsxPath}");
                return 1;
            }

            string outputPath = Path.ChangeExtension(xlsxPath, ".ct");

            Console.WriteLine($"Converting: {xlsxPath}");
            Console.WriteLine($"Output: {outputPath}");

            try
            {
                var processor = new CTProcessor();
                processor.ImportFromXLSX(xlsxPath);

                Console.WriteLine($"Columns: {processor.Headers.Count}");
                Console.WriteLine($"Rows: {processor.Rows.Count}");

                processor.Write(outputPath, processor.Headers, processor.Types, processor.Rows);

                Console.WriteLine($"Done!");
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return 1;
            }
        }

        static int PackDirectory(string[] args)
        {
            string sourceDir = args[1];

            if (!Directory.Exists(sourceDir))
            {
                Console.WriteLine($"Error: Directory not found: {sourceDir}");
                return 1;
            }

            // Determine output file
            string outputFile;
            if (args.Length >= 3 && !string.IsNullOrEmpty(args[2]))
            {
                outputFile = args[2];
            }
            else
            {
                // Default: directory name + .VDK
                string dirName = Path.GetFileName(sourceDir.TrimEnd('\\', '/'));
                string parentDir = Path.GetDirectoryName(sourceDir);
                outputFile = Path.Combine(parentDir, dirName + ".VDK");
            }

            Console.WriteLine($"Packing: {sourceDir}");
            Console.WriteLine($"Output: {outputFile}");

            try
            {
                var writer = new VDKWriter();

                Console.WriteLine("Reading files...");
                writer.AddDirectory(sourceDir, (current, total, file) =>
                {
                    if (current % 100 == 0 || current == total)
                        Console.WriteLine($"  Added {current}/{total} files...");
                });

                Console.WriteLine("Compressing and writing...");
                int count = writer.Write(outputFile, compress: true);

                Console.WriteLine();
                Console.WriteLine($"Done! Packed {count} files");
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return 1;
            }
        }
    }
}
