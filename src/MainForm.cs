using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VDKTool
{
    public class MainForm : Form
    {
        private MenuStrip menuStrip;
        private ToolStrip toolStrip;
        private SplitContainer splitContainer;
        private TreeView treeView;
        private ListView listView;
        private RichTextBox logBox;
        private StatusStrip statusStrip;
        private ToolStripStatusLabel statusLabel;
        private ToolStripProgressBar progressBar;
        private TextBox filterBox;

        private VDKArchive currentArchive;
        private ImageList imageList;

        public MainForm()
        {
            InitializeComponent();
            SetupImageList();
        }

        private void InitializeComponent()
        {
            this.Text = "VDK Tool - RO2 Archive Manager";
            this.Size = new Size(1000, 700);
            this.MinimumSize = new Size(800, 500);
            this.StartPosition = FormStartPosition.CenterScreen;

            // Menu
            menuStrip = new MenuStrip();

            var fileMenu = new ToolStripMenuItem("&File");
            fileMenu.DropDownItems.Add("&Open VDK...", null, (s, e) => OpenVDK());
            fileMenu.DropDownItems.Add("&Pack Directory...", null, (s, e) => PackDirectory());
            fileMenu.DropDownItems.Add(new ToolStripSeparator());
            fileMenu.DropDownItems.Add("E&xit", null, (s, e) => Close());
            menuStrip.Items.Add(fileMenu);

            var ctMenu = new ToolStripMenuItem("&CT Files");
            ctMenu.DropDownItems.Add("&Convert CT to CSV...", null, (s, e) => ConvertCTtoCSV());
            ctMenu.DropDownItems.Add("Convert &CSV to CT...", null, (s, e) => ConvertCSVtoCT());
            ctMenu.DropDownItems.Add(new ToolStripSeparator());
            ctMenu.DropDownItems.Add("&Batch Convert CT to CSV...", null, (s, e) => BatchConvertCT());
            ctMenu.DropDownItems.Add("Batch Convert CS&V to CT...", null, (s, e) => BatchConvertCSV());
            menuStrip.Items.Add(ctMenu);

            var toolsMenu = new ToolStripMenuItem("&Tools");
            toolsMenu.DropDownItems.Add("&Batch Extract VDK...", null, (s, e) => BatchExtract());
            toolsMenu.DropDownItems.Add("&Compare Archives...", null, (s, e) => CompareArchives());
            menuStrip.Items.Add(toolsMenu);

            var helpMenu = new ToolStripMenuItem("&Help");
            helpMenu.DropDownItems.Add("&About", null, (s, e) => ShowAbout());
            menuStrip.Items.Add(helpMenu);

            // Toolbar
            toolStrip = new ToolStrip();
            toolStrip.Items.Add(new ToolStripButton("Open", null, (s, e) => OpenVDK()) { DisplayStyle = ToolStripItemDisplayStyle.ImageAndText });
            toolStrip.Items.Add(new ToolStripButton("Extract All", null, (s, e) => ExtractAll()) { DisplayStyle = ToolStripItemDisplayStyle.ImageAndText });
            toolStrip.Items.Add(new ToolStripButton("Extract Selected", null, (s, e) => ExtractSelected()) { DisplayStyle = ToolStripItemDisplayStyle.ImageAndText });
            toolStrip.Items.Add(new ToolStripSeparator());
            toolStrip.Items.Add(new ToolStripButton("Pack", null, (s, e) => PackDirectory()) { DisplayStyle = ToolStripItemDisplayStyle.ImageAndText });
            toolStrip.Items.Add(new ToolStripSeparator());
            toolStrip.Items.Add(new ToolStripLabel("Filter:"));
            filterBox = new TextBox { Width = 150 };
            filterBox.TextChanged += FilterBox_TextChanged;
            toolStrip.Items.Add(new ToolStripControlHost(filterBox));

            // Split container
            splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 450
            };

            // Top split: Tree and List
            var topSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 250
            };

            // Tree view
            treeView = new TreeView
            {
                Dock = DockStyle.Fill,
                HideSelection = false
            };
            treeView.AfterSelect += TreeView_AfterSelect;
            topSplit.Panel1.Controls.Add(treeView);

            // List view
            listView = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true
            };
            listView.Columns.Add("Name", 250);
            listView.Columns.Add("Size", 100, HorizontalAlignment.Right);
            listView.Columns.Add("Compressed", 100, HorizontalAlignment.Right);
            listView.Columns.Add("Ratio", 70, HorizontalAlignment.Right);
            listView.DoubleClick += ListView_DoubleClick;
            listView.KeyDown += ListView_KeyDown;
            topSplit.Panel2.Controls.Add(listView);

            splitContainer.Panel1.Controls.Add(topSplit);

            // Log box
            var logLabel = new Label { Text = "Output Log", Dock = DockStyle.Top, Height = 20 };
            logBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = Color.White,
                Font = new Font("Consolas", 9)
            };
            splitContainer.Panel2.Controls.Add(logBox);
            splitContainer.Panel2.Controls.Add(logLabel);

            // Status bar
            statusStrip = new StatusStrip();
            statusLabel = new ToolStripStatusLabel("Ready") { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
            progressBar = new ToolStripProgressBar { Width = 200, Visible = false };
            statusStrip.Items.Add(statusLabel);
            statusStrip.Items.Add(progressBar);

            // Layout
            this.MainMenuStrip = menuStrip;
            this.Controls.Add(splitContainer);
            this.Controls.Add(toolStrip);
            this.Controls.Add(menuStrip);
            this.Controls.Add(statusStrip);
        }

        private void SetupImageList()
        {
            imageList = new ImageList { ImageSize = new Size(16, 16) };
            treeView.ImageList = imageList;
            listView.SmallImageList = imageList;
        }

        private void Log(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => Log(message)));
                return;
            }
            logBox.AppendText(message + Environment.NewLine);
            logBox.ScrollToCaret();
        }

        private void SetStatus(string text)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => SetStatus(text)));
                return;
            }
            statusLabel.Text = text;
        }

        private void SetProgress(int value, int max)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => SetProgress(value, max)));
                return;
            }
            progressBar.Maximum = max;
            progressBar.Value = value;
            progressBar.Visible = max > 0;
        }

        private void OpenVDK()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "VDK files (*.vdk)|*.vdk|All files (*.*)|*.*";
                dialog.Title = "Open VDK Archive";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    LoadArchive(dialog.FileName);
                }
            }
        }

        private void LoadArchive(string filePath)
        {
            try
            {
                SetStatus($"Loading {Path.GetFileName(filePath)}...");
                Cursor = Cursors.WaitCursor;

                currentArchive = VDKArchive.Load(filePath);

                var fileEntries = currentArchive.GetFileEntries();
                long totalUncomp = fileEntries.Sum(e => e.UncompressedSize);
                long totalComp = fileEntries.Sum(e => e.CompressedSize);
                double ratio = totalUncomp > 0 ? (double)totalComp / totalUncomp * 100 : 0;

                Log($"Loaded: {filePath}");
                Log($"  Format: {currentArchive.Version}");
                Log($"  Files: {currentArchive.FileCount}, Folders: {currentArchive.FolderCount}");
                Log($"  Size: {totalUncomp:N0} bytes ({ratio:F1}% compressed)");

                PopulateTree();
                SetStatus($"{Path.GetFileName(filePath)} - {currentArchive.FileCount} files");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load archive:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                SetStatus("Ready");
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void PopulateTree()
        {
            treeView.Nodes.Clear();
            listView.Items.Clear();

            if (currentArchive == null) return;

            string filter = filterBox.Text.ToLowerInvariant();
            var entries = currentArchive.Entries.Where(e => e.Name != "." && e.Name != "..").ToList();

            if (!string.IsNullOrEmpty(filter))
            {
                entries = entries.Where(e =>
                    e.Path.ToLowerInvariant().Contains(filter) ||
                    e.Name.ToLowerInvariant().Contains(filter)).ToList();
            }

            // Build tree
            var nodes = new Dictionary<string, TreeNode>();
            var rootNode = new TreeNode(Path.GetFileName(currentArchive.FilePath)) { Tag = "" };
            treeView.Nodes.Add(rootNode);
            nodes[""] = rootNode;

            foreach (var entry in entries.Where(e => e.IsDirectory).OrderBy(e => e.Path))
            {
                string parentPath = Path.GetDirectoryName(entry.Path)?.Replace('\\', '/') ?? "";
                TreeNode parentNode = nodes.ContainsKey(parentPath) ? nodes[parentPath] : rootNode;

                var node = new TreeNode(entry.Name) { Tag = entry.Path };
                parentNode.Nodes.Add(node);
                nodes[entry.Path] = node;
            }

            rootNode.Expand();

            // Select root to show files
            treeView.SelectedNode = rootNode;
        }

        private void TreeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (currentArchive == null || e.Node == null) return;

            string selectedPath = e.Node.Tag as string ?? "";
            PopulateListView(selectedPath);
        }

        private void PopulateListView(string folderPath)
        {
            listView.Items.Clear();

            string filter = filterBox.Text.ToLowerInvariant();
            IEnumerable<FileEntry> entries = currentArchive.Entries
                .Where(e => e.Name != "." && e.Name != ".." && !e.IsDirectory)
                .Where(e =>
                {
                    string entryFolder = Path.GetDirectoryName(e.Path)?.Replace('\\', '/') ?? "";
                    return entryFolder.Equals(folderPath, StringComparison.OrdinalIgnoreCase) ||
                           (string.IsNullOrEmpty(folderPath) && string.IsNullOrEmpty(entryFolder));
                })
                .OrderBy(e => e.Name);

            if (!string.IsNullOrEmpty(filter))
            {
                entries = entries.Where(e => e.Name.ToLowerInvariant().Contains(filter));
            }

            foreach (var entry in entries)
            {
                double ratio = entry.UncompressedSize > 0 ? (double)entry.CompressedSize / entry.UncompressedSize * 100 : 0;
                var item = new ListViewItem(entry.Name);
                item.SubItems.Add($"{entry.UncompressedSize:N0}");
                item.SubItems.Add($"{entry.CompressedSize:N0}");
                item.SubItems.Add($"{ratio:F1}%");
                item.Tag = entry;
                listView.Items.Add(item);
            }

            SetStatus($"{listView.Items.Count} files in folder");
        }

        private void FilterBox_TextChanged(object sender, EventArgs e)
        {
            if (currentArchive != null)
            {
                PopulateTree();
            }
        }

        private void ListView_DoubleClick(object sender, EventArgs e)
        {
            ExtractSelected();
        }

        private void ListView_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ExtractSelected();
            }
        }

        private void ExtractAll()
        {
            if (currentArchive == null)
            {
                MessageBox.Show("No archive loaded", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select output directory";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var entries = currentArchive.Entries.Where(e => e.Name != "." && e.Name != "..").ToList();
                    ExtractEntries(entries, dialog.SelectedPath);
                }
            }
        }

        private void ExtractSelected()
        {
            if (currentArchive == null)
            {
                MessageBox.Show("No archive loaded", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var selectedEntries = new List<FileEntry>();

            foreach (ListViewItem item in listView.SelectedItems)
            {
                if (item.Tag is FileEntry entry)
                    selectedEntries.Add(entry);
            }

            if (selectedEntries.Count == 0)
            {
                MessageBox.Show("No files selected", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select output directory";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    ExtractEntries(selectedEntries, dialog.SelectedPath);
                }
            }
        }

        private async void ExtractEntries(List<FileEntry> entries, string outputDir)
        {
            var fileEntries = entries.Where(e => !e.IsDirectory).ToList();
            int total = fileEntries.Count;
            int extracted = 0;

            SetProgress(0, total);
            SetStatus("Extracting...");

            await Task.Run(() =>
            {
                foreach (var entry in entries)
                {
                    if (entry.Name == "." || entry.Name == "..") continue;

                    if (entry.IsDirectory)
                    {
                        string dirPath = Path.Combine(outputDir, entry.Path);
                        Directory.CreateDirectory(dirPath);
                        continue;
                    }

                    try
                    {
                        string outPath = Path.Combine(outputDir, entry.Path);
                        Directory.CreateDirectory(Path.GetDirectoryName(outPath));

                        byte[] data = currentArchive.ExtractFile(entry);
                        File.WriteAllBytes(outPath, data);

                        extracted++;
                        SetProgress(extracted, total);
                        SetStatus($"Extracting: {entry.Name}");
                    }
                    catch (Exception ex)
                    {
                        Log($"ERROR: {entry.Path} - {ex.Message}");
                    }
                }
            });

            SetProgress(0, 0);
            Log($"Extracted {extracted} files to {outputDir}");
            SetStatus($"Extracted {extracted} files");
        }

        private void PackDirectory()
        {
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select directory to pack";
                if (folderDialog.ShowDialog() != DialogResult.OK) return;

                string sourceDir = folderDialog.SelectedPath;

                using (var saveDialog = new SaveFileDialog())
                {
                    saveDialog.Filter = "VDK files (*.vdk)|*.vdk";
                    saveDialog.FileName = Path.GetFileName(sourceDir) + ".vdk";

                    if (saveDialog.ShowDialog() != DialogResult.OK) return;

                    PackDirectoryAsync(sourceDir, saveDialog.FileName);
                }
            }
        }

        private async void PackDirectoryAsync(string sourceDir, string outputFile)
        {
            SetStatus("Packing...");
            Log($"Packing: {sourceDir}");

            int count = 0;

            await Task.Run(() =>
            {
                var writer = new VDKWriter();

                writer.AddDirectory(sourceDir, (current, total, file) =>
                {
                    SetProgress(current, total);
                    SetStatus($"Adding: {file}");
                });

                Log("Compressing and writing...");
                count = writer.Write(outputFile, compress: true);
            });

            long origSize = Directory.GetFiles(sourceDir, "*", SearchOption.AllDirectories).Sum(f => new FileInfo(f).Length);
            long archSize = new FileInfo(outputFile).Length;
            double ratio = origSize > 0 ? (double)archSize / origSize * 100 : 0;

            SetProgress(0, 0);
            Log($"Packed {count} files");
            Log($"Original: {origSize:N0} bytes");
            Log($"Archive: {archSize:N0} bytes ({ratio:F1}%)");
            SetStatus($"Packed {count} files");
        }

        private void BatchExtract()
        {
            using (var inputDialog = new FolderBrowserDialog())
            {
                inputDialog.Description = "Select directory containing VDK files";
                if (inputDialog.ShowDialog() != DialogResult.OK) return;

                using (var outputDialog = new FolderBrowserDialog())
                {
                    outputDialog.Description = "Select output directory";
                    if (outputDialog.ShowDialog() != DialogResult.OK) return;

                    BatchExtractAsync(inputDialog.SelectedPath, outputDialog.SelectedPath);
                }
            }
        }

        private async void BatchExtractAsync(string inputDir, string outputDir)
        {
            var vdkFiles = Directory.GetFiles(inputDir, "*.vdk", SearchOption.AllDirectories);

            Log($"Found {vdkFiles.Length} VDK files");
            SetStatus("Batch extracting...");

            int totalFiles = 0;

            await Task.Run(() =>
            {
                for (int i = 0; i < vdkFiles.Length; i++)
                {
                    string vdkPath = vdkFiles[i];
                    string relPath = vdkPath.Substring(inputDir.Length).TrimStart('\\', '/');
                    string vdkOutput = Path.Combine(outputDir, Path.ChangeExtension(relPath, null));

                    SetProgress(i + 1, vdkFiles.Length);
                    SetStatus($"Processing: {relPath}");

                    try
                    {
                        var archive = VDKArchive.Load(vdkPath);

                        foreach (var entry in archive.Entries)
                        {
                            if (entry.Name == "." || entry.Name == "..") continue;

                            if (entry.IsDirectory)
                            {
                                Directory.CreateDirectory(Path.Combine(vdkOutput, entry.Path));
                                continue;
                            }

                            string outPath = Path.Combine(vdkOutput, entry.Path);
                            Directory.CreateDirectory(Path.GetDirectoryName(outPath));

                            byte[] data = archive.ExtractFile(entry);
                            File.WriteAllBytes(outPath, data);
                            totalFiles++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"ERROR: {relPath} - {ex.Message}");
                    }
                }
            });

            SetProgress(0, 0);
            Log($"Batch extraction complete: {totalFiles} files from {vdkFiles.Length} archives");
            SetStatus($"Extracted {totalFiles} files");
        }

        private void CompareArchives()
        {
            using (var dialog1 = new OpenFileDialog())
            {
                dialog1.Filter = "VDK files (*.vdk)|*.vdk|All files (*.*)|*.*";
                dialog1.Title = "Select first archive";

                if (dialog1.ShowDialog() != DialogResult.OK) return;

                using (var dialog2 = new OpenFileDialog())
                {
                    dialog2.Filter = "VDK files (*.vdk)|*.vdk|All files (*.*)|*.*";
                    dialog2.Title = "Select second archive";

                    if (dialog2.ShowDialog() != DialogResult.OK) return;

                    CompareArchivesAsync(dialog1.FileName, dialog2.FileName);
                }
            }
        }

        private async void CompareArchivesAsync(string file1, string file2)
        {
            SetStatus("Comparing...");
            Log($"Comparing: {Path.GetFileName(file1)} vs {Path.GetFileName(file2)}");

            await Task.Run(() =>
            {
                try
                {
                    var archive1 = VDKArchive.Load(file1);
                    var archive2 = VDKArchive.Load(file2);

                    var files1 = archive1.GetFileEntries().ToDictionary(e => e.Path.ToLowerInvariant(), e => e);
                    var files2 = archive2.GetFileEntries().ToDictionary(e => e.Path.ToLowerInvariant(), e => e);

                    var allPaths = new HashSet<string>(files1.Keys);
                    allPaths.UnionWith(files2.Keys);

                    var onlyIn1 = new List<string>();
                    var onlyIn2 = new List<string>();
                    var different = new List<string>();
                    var identical = new List<string>();

                    foreach (var path in allPaths.OrderBy(p => p))
                    {
                        bool in1 = files1.ContainsKey(path);
                        bool in2 = files2.ContainsKey(path);

                        if (in1 && !in2) onlyIn1.Add(path);
                        else if (in2 && !in1) onlyIn2.Add(path);
                        else if (files1[path].UncompressedSize != files2[path].UncompressedSize)
                            different.Add(path);
                        else identical.Add(path);
                    }

                    Log("");
                    Log($"Identical: {identical.Count} files");
                    Log($"Different: {different.Count} files");
                    Log($"Only in first: {onlyIn1.Count} files");
                    Log($"Only in second: {onlyIn2.Count} files");

                    if (different.Count > 0)
                    {
                        Log("\nDifferent files:");
                        foreach (var p in different.Take(10))
                            Log($"  {p}");
                        if (different.Count > 10)
                            Log($"  ... and {different.Count - 10} more");
                    }
                }
                catch (Exception ex)
                {
                    Log($"ERROR: {ex.Message}");
                }
            });

            SetStatus("Comparison complete");
        }

        private void ConvertCTtoCSV()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "CT files (*.ct)|*.ct|All files (*.*)|*.*";
                dialog.Title = "Select CT file to convert";
                dialog.Multiselect = true;

                if (dialog.ShowDialog() != DialogResult.OK) return;

                foreach (var file in dialog.FileNames)
                {
                    try
                    {
                        var processor = new CTProcessor();
                        processor.Read(file);

                        string outputPath = Path.ChangeExtension(file, ".csv");
                        processor.ExportToCSV(outputPath);

                        Log($"Converted: {Path.GetFileName(file)} -> {Path.GetFileName(outputPath)}");
                    }
                    catch (Exception ex)
                    {
                        Log($"ERROR converting {Path.GetFileName(file)}: {ex.Message}");
                    }
                }

                SetStatus($"Converted {dialog.FileNames.Length} files");
            }
        }

        private void ConvertCSVtoCT()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                dialog.Title = "Select CSV file to convert";
                dialog.Multiselect = true;

                if (dialog.ShowDialog() != DialogResult.OK) return;

                foreach (var file in dialog.FileNames)
                {
                    try
                    {
                        var processor = new CTProcessor();
                        processor.ImportFromCSV(file);

                        string outputPath = Path.ChangeExtension(file, ".ct");
                        processor.Write(outputPath, processor.Headers, processor.Types, processor.Rows);

                        Log($"Converted: {Path.GetFileName(file)} -> {Path.GetFileName(outputPath)}");
                    }
                    catch (Exception ex)
                    {
                        Log($"ERROR converting {Path.GetFileName(file)}: {ex.Message}");
                    }
                }

                SetStatus($"Converted {dialog.FileNames.Length} files");
            }
        }

        private async void BatchConvertCT()
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select folder containing CT files";
                if (dialog.ShowDialog() != DialogResult.OK) return;

                string inputDir = dialog.SelectedPath;
                var ctFiles = Directory.GetFiles(inputDir, "*.ct", SearchOption.AllDirectories);

                Log($"Found {ctFiles.Length} CT files");
                SetStatus("Converting CT files...");

                int converted = 0;
                int errors = 0;

                await Task.Run(() =>
                {
                    for (int i = 0; i < ctFiles.Length; i++)
                    {
                        SetProgress(i + 1, ctFiles.Length);
                        SetStatus($"Converting: {Path.GetFileName(ctFiles[i])}");

                        try
                        {
                            var processor = new CTProcessor();
                            processor.Read(ctFiles[i]);

                            string outputPath = Path.ChangeExtension(ctFiles[i], ".csv");
                            processor.ExportToCSV(outputPath);
                            converted++;
                        }
                        catch (Exception ex)
                        {
                            Log($"ERROR: {Path.GetFileName(ctFiles[i])} - {ex.Message}");
                            errors++;
                        }
                    }
                });

                SetProgress(0, 0);
                Log($"Batch conversion complete: {converted} converted, {errors} errors");
                SetStatus($"Converted {converted} CT files");
            }
        }

        private async void BatchConvertCSV()
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select folder containing CSV files";
                if (dialog.ShowDialog() != DialogResult.OK) return;

                string inputDir = dialog.SelectedPath;
                var csvFiles = Directory.GetFiles(inputDir, "*.csv", SearchOption.AllDirectories);

                Log($"Found {csvFiles.Length} CSV files");
                SetStatus("Converting CSV files...");

                int converted = 0;
                int errors = 0;

                await Task.Run(() =>
                {
                    for (int i = 0; i < csvFiles.Length; i++)
                    {
                        SetProgress(i + 1, csvFiles.Length);
                        SetStatus($"Converting: {Path.GetFileName(csvFiles[i])}");

                        try
                        {
                            var processor = new CTProcessor();
                            processor.ImportFromCSV(csvFiles[i]);

                            string outputPath = Path.ChangeExtension(csvFiles[i], ".ct");
                            processor.Write(outputPath, processor.Headers, processor.Types, processor.Rows);
                            converted++;
                        }
                        catch (Exception ex)
                        {
                            Log($"ERROR: {Path.GetFileName(csvFiles[i])} - {ex.Message}");
                            errors++;
                        }
                    }
                });

                SetProgress(0, 0);
                Log($"Batch conversion complete: {converted} converted, {errors} errors");
                SetStatus($"Converted {converted} CSV files");
            }
        }

        private void ShowAbout()
        {
            MessageBox.Show(
                "VDK Tool - RO2 Archive Manager\n\n" +
                "A tool for working with Ragnarok Online 2 archives.\n\n" +
                "Features:\n" +
                "- VDK archive extraction and creation\n" +
                "- CT table file conversion (CT <-> CSV)\n" +
                "- Batch processing support\n\n" +
                "Supports VDISK1.0 and VDISK1.1 formats.\n\n" +
                "Author: JordanRO2",
                "About VDK Tool",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
    }
}
