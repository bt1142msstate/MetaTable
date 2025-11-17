using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.ComponentModel;
using OfficeOpenXml;
using System.Security.Cryptography;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using System.Windows.Threading;

// Add aliases to resolve ambiguity
using WpfMessageBox = System.Windows.MessageBox;
using WpfApplication = System.Windows.Application;
using WinFormsApplication = System.Windows.Forms.Application;
using IODirectory = System.IO.Directory;

namespace MetaTable
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {        private ObservableCollection<FileMetadata> _fileMetadataList;
        private BackgroundWorker _scanWorker;
        private string _selectedFolderPath = "";
        private FileSystemWatcher _fileWatcher;
        private DispatcherTimer _fileStatusTimer;
        private bool _isExcelFileInUse = false;        // Animation properties
        private DispatcherTimer _animationTimer;
        private Queue<string> _fileQueue = new Queue<string>();
        private Border _currentFileCard;
        private Random _random = new Random();
        private System.Windows.Media.Animation.Storyboard _scanningStoryboard;
        
        // File icon cycling for animation
        private string[] _fileIcons = { "📄", "📁", "📊", "🖼️", "📝", "🔧", "📋", "📈", "🎵", "🎬" };
        private int _currentIconIndex = 0;
        private DispatcherTimer _iconCycleTimer;
        
        // Realistic metadata samples for animation
        private readonly string[][] _metadataSamples = new string[][]
        {
            new string[] { "📁 Size: 2.4 MB", "📅 Modified: 2025-06-21", "🏷️ Type: PDF Document", "👤 Author: John Smith", "📐 Pages: 15" },
            new string[] { "📁 Size: 847 KB", "📅 Modified: 2025-06-20", "🏷️ Type: Excel File", "👤 Author: Sarah Johnson", "📐 Sheets: 3" },
            new string[] { "📁 Size: 12.8 MB", "📅 Modified: 2025-06-19", "🏷️ Type: Image (JPEG)", "📷 Camera: Canon EOS R5", "📐 Dimensions: 6000x4000" },
            new string[] { "📁 Size: 156 KB", "📅 Modified: 2025-06-18", "🏷️ Type: Word Document", "👤 Author: Mike Wilson", "📝 Words: 2,847" },
            new string[] { "📁 Size: 4.7 MB", "📅 Modified: 2025-06-17", "🏷️ Type: Audio (MP3)", "🎵 Artist: The Beatles", "⏱️ Duration: 3:24" },
            new string[] { "📁 Size: 892 KB", "📅 Modified: 2025-06-16", "🏷️ Type: PowerPoint", "👤 Author: Lisa Chen", "📊 Slides: 28" },
            new string[] { "📁 Size: 67.2 MB", "📅 Modified: 2025-06-15", "🏷️ Type: Video (MP4)", "🎬 Codec: H.264", "⏱️ Duration: 5:12" },
            new string[] { "📁 Size: 234 KB", "📅 Modified: 2025-06-14", "🏷️ Type: Text File", "📝 Encoding: UTF-8", "📐 Lines: 1,456" }
        };
        private int _currentMetadataIndex = 0;public MainWindow()
        {
            InitializeComponent();
            _fileMetadataList = new ObservableCollection<FileMetadata>();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;              // Initialize scanning animation
            _scanningStoryboard = (System.Windows.Media.Animation.Storyboard)this.Resources["ScanningAnimation"];
              // Initialize icon cycling timer for animation
            _iconCycleTimer = new DispatcherTimer();
            _iconCycleTimer.Interval = TimeSpan.FromSeconds(4); // Match animation duration
            _iconCycleTimer.Tick += IconCycleTimer_Tick;
            
            OutputFileTextBox.Text = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                $"FileMetadata_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");            FolderPathTextBox.Text = "📁 No folder selected - Click 'Select Folder to Scan' to choose a directory";
            StatusBarText.Text = "Ready • Select a folder to begin scanning";
            StatusTextBlock.Text = "Ready to scan";
            ScanButton.IsEnabled = false;
            ScanButton.Visibility = Visibility.Hidden;
              // Initialize file monitoring
            InitializeFileMonitoring();
        }private void SelectFolderButton_Click(object sender, RoutedEventArgs e)
        {
            using (var folderDialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                folderDialog.Description = "Select Folder to Scan for File Metadata";
                folderDialog.ShowNewFolderButton = false;

                // Set initial directory if current path is valid
                if (!string.IsNullOrEmpty(FolderPathTextBox.Text) && IODirectory.Exists(FolderPathTextBox.Text))
                {
                    folderDialog.SelectedPath = FolderPathTextBox.Text;
                }
                else
                {
                    folderDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                }

                if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    _selectedFolderPath = folderDialog.SelectedPath;
                    FolderPathTextBox.Text = _selectedFolderPath;
                
                // Validate the selected folder and enable scan button
                if (IODirectory.Exists(_selectedFolderPath))
                {
                    if (ScanButton != null)
                    {
                        ScanButton.Visibility = Visibility.Visible;

                        // Check if Excel file is currently in use before enabling scan button
                        string outputPath = OutputFileTextBox.Text;
                        bool isExcelFileInUse = !string.IsNullOrWhiteSpace(outputPath) && IsFileInUse(outputPath);                        if (isExcelFileInUse)
                        {
                            // Excel file is open - keep scan button disabled
                            ScanButton.IsEnabled = false;
                            ScanButton.Content = "Excel File Open - Close it to enable scanning";
                            ScanButton.Opacity = 0.6;
                            string fileName = Path.GetFileName(outputPath);
                            StatusBarText.Text = $"📁 Folder selected but Excel file '{fileName}' is open - Close it to enable scanning";
                            StatusTextBlock.Text = "📄 Excel file is open";
                        }                        else
                        {
                            // Excel file is available - enable scanning
                            ScanButton.IsEnabled = true;
                            ScanButton.Content = "Scan Files and Export Data to Excel";
                            ScanButton.Opacity = 1.0;
                            
                            // Set tooltip for scan mode
                            var tooltip = ScanButton.ToolTip as System.Windows.Controls.ToolTip;
                            if (tooltip != null)
                            {
                                tooltip.Content = "🚀 Start scanning selected folder and export metadata to Excel file";
                            }
                            
                            StatusBarText.Text = $"📁 Selected: {System.IO.Path.GetFileName(_selectedFolderPath)} • Ready to scan";
                            StatusTextBlock.Text = "📂 Folder selected";
                        }
                    }
                }
                else
                {
                    WpfMessageBox.Show("Selected folder is not valid or accessible.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    _selectedFolderPath = "";
                    if (ScanButton != null)
                    {
                        ScanButton.IsEnabled = false;
                        ScanButton.Visibility = Visibility.Hidden;
                    }                    StatusBarText.Text = "❌ Please select a valid folder";
                }
                _fileMetadataList.Clear();
                // Remove ExportButton logic
                }
            }
        }private void FolderPathTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            // Reset scan button when folder path changes manually (not through folder picker)
            if (ScanButton != null && FolderPathTextBox.Text != _selectedFolderPath)
            {
                ScanButton.IsEnabled = false;
                ScanButton.Visibility = Visibility.Hidden;
                _selectedFolderPath = "";
                StatusBarText.Text = "📁 Please select a folder using the Browse button";
                StatusTextBlock.Text = "Ready to scan";
            }
        }        private void IconCycleTimer_Tick(object sender, EventArgs e)
        {
            // Only cycle through different file icons during animation (metadata now comes from real files)
            _currentIconIndex = (_currentIconIndex + 1) % _fileIcons.Length;
            
            // Don't update the file icon text here anymore - it's updated with real metadata
            // ScanningFileIcon.Text will be updated by UpdateAnimationWithRealMetadata()
        }        private void ScanButton_Click(object sender, RoutedEventArgs e)
        {
            // Check if we're currently scanning - if so, cancel the scan
            if (_scanWorker?.IsBusy == true)
            {
                _scanWorker.CancelAsync();
                ScanButton.Content = "Cancelling...";
                ScanButton.IsEnabled = false;
                  // Update tooltip for cancelling state
                var cancelTooltip = ScanButton.ToolTip as System.Windows.Controls.ToolTip;
                if (cancelTooltip != null)
                {
                    cancelTooltip.Content = "⏹️ Cancelling scan operation, please wait...";
                }
                
                StatusTextBlock.Text = "Cancelling scan...";
                StatusBarText.Text = "Cancelling scan, please wait...";
                return;
            }
            
            if (string.IsNullOrEmpty(_selectedFolderPath) || !IODirectory.Exists(_selectedFolderPath))
            {
                WpfMessageBox.Show("Please select a valid folder first.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Check if Excel file is currently in use before starting scan
            string outputPath = OutputFileTextBox.Text;
            if (!string.IsNullOrWhiteSpace(outputPath) && IsFileInUse(outputPath))
            {
                string fileName = Path.GetFileName(outputPath);
                WpfMessageBox.Show($"❌ Cannot start scan because the Excel file is currently open!\n\n" +
                                 $"📁 File: {fileName}\n\n" +
                                 $"💡 Solution: Please close the Excel file and try again.\n\n" +
                                 $"The file might be open in:\n" +
                                 $"• Microsoft Excel\n" +
                                 $"• Another spreadsheet application\n" +
                                 $"• File preview in Windows Explorer", 
                                 "Excel File Is Open", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }            _fileMetadataList.Clear();
            ScanButton.IsEnabled = true;
            ScanButton.Content = "⏹️ Cancel Scan";
            ScanButton.Opacity = 1.0;
              // Update tooltip for cancel mode
            var scanTooltip = ScanButton.ToolTip as System.Windows.Controls.ToolTip;
            if (scanTooltip != null)
            {
                scanTooltip.Content = "⏹️ Click to cancel the current scan operation";
            }
            
            // Disable Select Folder and Select Output File buttons while scanning
            SelectFolderButton.IsEnabled = false;
            BrowseOutputButton.IsEnabled = false;
            
            // Start scanning animation
            ScanAnimation.Visibility = Visibility.Visible;
            _scanningStoryboard?.Begin();
            _iconCycleTimer?.Start();
            
            // Remove ExportButton logic
            _scanWorker = new BackgroundWorker
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };
            
            _scanWorker.DoWork += ScanWorker_DoWork;
            _scanWorker.ProgressChanged += ScanWorker_ProgressChanged;
            _scanWorker.RunWorkerCompleted += ScanWorker_RunWorkerCompleted;
            
            _scanWorker.RunWorkerAsync();
        }        private void ScanWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;
            var files = new List<FileMetadata>();
            
            try
            {
                var allFiles = IODirectory.GetFiles(_selectedFolderPath, "*", SearchOption.AllDirectories);
                int totalFiles = allFiles.Length;
                
                // Get UI settings once for all processing
                bool includeHidden = false;
                bool includeSystem = false;
                
                WpfApplication.Current.Dispatcher.Invoke(() =>
                {
                    includeHidden = IncludeHiddenFilesCheckBox.IsChecked == true;
                    includeSystem = IncludeSystemFilesCheckBox.IsChecked == true;
                });

                // Use parallel processing for large scans (1000+ files) to improve performance
                if (totalFiles >= 1000)
                {
                    files = ProcessFilesInParallel(allFiles, worker, includeHidden, includeSystem, e);
                }
                else
                {
                    files = ProcessFilesSequentially(allFiles, worker, includeHidden, includeSystem);
                }

                if (worker?.CancellationPending == true)
                {
                    e.Cancel = true;
                    return;
                }

                e.Result = files;
            }
            catch (Exception ex)
            {
                e.Result = ex;
            }
        }

        private List<FileMetadata> ProcessFilesSequentially(string[] allFiles, BackgroundWorker worker, bool includeHidden, bool includeSystem)
        {
            var files = new List<FileMetadata>();
            int totalFiles = allFiles.Length;
            int processedFiles = 0;

            foreach (var filePath in allFiles)
            {
                if (worker?.CancellationPending == true)
                {
                    break;
                }

                try
                {
                    // Quick attribute check to filter files early
                    var attributes = File.GetAttributes(filePath);
                    
                    if (!includeHidden && (attributes & FileAttributes.Hidden) == FileAttributes.Hidden)
                        continue;
                    if (!includeSystem && (attributes & FileAttributes.System) == FileAttributes.System)
                        continue;

                    // Extract comprehensive metadata (always use non-diagnostic mode)
                    var metadata = ExtractComprehensiveMetadata(filePath, false);

                    files.Add(metadata);
                    processedFiles++;

                    // Adaptive progress reporting based on total file count
                    int updateInterval = totalFiles switch
                    {
                        <= 50 => 1,      // Update every file for small scans
                        <= 200 => 2,     // Update every 2 files for small-medium scans
                        <= 500 => 5,     // Update every 5 files for medium scans
                        <= 2000 => 10,   // Update every 10 files for large scans
                        _ => 25          // Update every 25 files for very large scans
                    };                    if (processedFiles % updateInterval == 0 || processedFiles == totalFiles)
                    {
                        worker?.ReportProgress((int)((double)processedFiles / totalFiles * 100), 
                                             new ProgressInfo { 
                                                 FilesProcessed = processedFiles, 
                                                 TotalFiles = totalFiles, 
                                                 CurrentFile = Path.GetFileName(filePath),
                                                 FilesFound = files.Count,
                                                 CurrentMetadata = metadata
                                             });
                    }
                }
                catch (Exception ex)
                {
                    // Skip files that can't be accessed
                    System.Diagnostics.Debug.WriteLine($"Error processing file {filePath}: {ex.Message}");
                }
            }

            return files;
        }

        private List<FileMetadata> ProcessFilesInParallel(string[] allFiles, BackgroundWorker worker, bool includeHidden, bool includeSystem, DoWorkEventArgs e)
        {
            var files = new ConcurrentBag<FileMetadata>();
            int totalFiles = allFiles.Length;
            int processedFiles = 0;
            var lockObject = new object();
            
            // Determine optimal parallelism based on system capabilities and file count
            int maxDegreeOfParallelism = Math.Min(Environment.ProcessorCount * 2, Math.Max(4, totalFiles / 1000));
            
            var parallelOptions = new ParallelOptions
            {
                MaxDegreeOfParallelism = maxDegreeOfParallelism,
                CancellationToken = System.Threading.CancellationToken.None
            };

            try
            {
                Parallel.ForEach(allFiles, parallelOptions, (filePath, loopState) =>
                {
                    if (worker?.CancellationPending == true)
                    {
                        loopState.Stop();
                        return;
                    }

                    try
                    {
                        // Quick attribute check to filter files early
                        var attributes = File.GetAttributes(filePath);
                        
                        if (!includeHidden && (attributes & FileAttributes.Hidden) == FileAttributes.Hidden)
                            return;
                        if (!includeSystem && (attributes & FileAttributes.System) == FileAttributes.System)
                            return;

                        // Extract comprehensive metadata (always use non-diagnostic mode)
                        var metadata = ExtractComprehensiveMetadata(filePath, false);
                        files.Add(metadata);

                        // Thread-safe progress reporting
                        lock (lockObject)
                        {
                            processedFiles++;
                            
                            // Adaptive progress reporting for parallel processing
                            int updateInterval = totalFiles switch
                            {
                                <= 1000 => 50,   // Update every 50 files for 1K files
                                <= 5000 => 100,  // Update every 100 files for 5K files
                                <= 10000 => 200, // Update every 200 files for 10K files
                                _ => 500         // Update every 500 files for very large scans
                            };                            if (processedFiles % updateInterval == 0 || processedFiles == totalFiles)
                            {
                                worker?.ReportProgress((int)((double)processedFiles / totalFiles * 100), 
                                                     new ProgressInfo { 
                                                         FilesProcessed = processedFiles, 
                                                         TotalFiles = totalFiles, 
                                                         CurrentFile = Path.GetFileName(filePath),
                                                         FilesFound = files.Count,
                                                         CurrentMetadata = metadata
                                                     });
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Skip files that can't be accessed
                        System.Diagnostics.Debug.WriteLine($"Error processing file {filePath}: {ex.Message}");
                    }
                });

                if (worker?.CancellationPending == true)
                {
                    e.Cancel = true;
                }
            }
            catch (OperationCanceledException)
            {
                e.Cancel = true;
            }

            return files.ToList();
        }        private void ScanWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            ProgressBar.Value = e.ProgressPercentage;
            if (e.UserState is ProgressInfo progress)
            {
                StatusTextBlock.Text = $"Scanning... {e.ProgressPercentage}% | {progress.FilesProcessed}/{progress.TotalFiles}";
                StatusBarText.Text = $"Processing: {progress.CurrentFile}";
                
                // Update animation with real metadata
                if (progress.CurrentMetadata != null)
                {
                    UpdateAnimationWithRealMetadata(progress.CurrentMetadata);
                }
            }
            else
            {
                // Fallback for compatibility
                StatusTextBlock.Text = $"Scanning... {e.ProgressPercentage}%";
            }
        }
        
        private void UpdateAnimationWithRealMetadata(FileMetadata metadata)
        {
            // Update file icon based on file type
            var extension = metadata.Extension.ToLower();
            string icon = extension switch
            {
                ".pdf" => "📄",
                ".doc" or ".docx" => "📝", 
                ".xls" or ".xlsx" => "📊",
                ".ppt" or ".pptx" => "📋",
                ".jpg" or ".jpeg" or ".png" or ".gif" or ".bmp" => "🖼️",
                ".mp3" or ".wav" or ".flac" => "🎵",
                ".mp4" or ".avi" or ".mkv" or ".mov" => "🎬",
                ".zip" or ".rar" or ".7z" => "📦",
                ".exe" or ".msi" => "🔧",
                ".txt" or ".log" => "📃",
                _ => "📄"
            };
            ScanningFileIcon.Text = icon;
            
            // Update metadata lines with real data
            MetadataLine1.Text = $"📁 Size: {FormatFileSize(metadata.SizeBytes)}";
            MetadataLine2.Text = $"📅 Modified: {metadata.ModifiedDate:yyyy-MM-dd}";
            MetadataLine3.Text = $"🏷️ Type: {GetFileTypeDescription(metadata.Extension)}";
            
            // Line 4: Try to get meaningful metadata based on file type
            string line4 = "👤 Author: Unknown";
            if (!string.IsNullOrEmpty(metadata.DocumentAuthor))
                line4 = $"👤 Author: {metadata.DocumentAuthor}";
            else if (!string.IsNullOrEmpty(metadata.Owner))
                line4 = $"👤 Owner: {metadata.Owner}";
            else if (!string.IsNullOrEmpty(metadata.Artist))
                line4 = $"🎵 Artist: {metadata.Artist}";
            else if (!string.IsNullOrEmpty(metadata.CameraMake))
                line4 = $"📷 Camera: {metadata.CameraMake}";
            MetadataLine4.Text = line4;
            
            // Line 5: Additional file-specific metadata
            string line5 = $"📐 Extension: {metadata.Extension}";
            if (!string.IsNullOrEmpty(metadata.ImageWidth) && !string.IsNullOrEmpty(metadata.ImageHeight))
                line5 = $"📐 Dimensions: {metadata.ImageWidth}x{metadata.ImageHeight}";
            else if (metadata.PageCount > 0)
                line5 = $"📄 Pages: {metadata.PageCount}";
            else if (metadata.WordCount > 0)
                line5 = $"📝 Words: {metadata.WordCount:N0}";
            else if (!string.IsNullOrEmpty(metadata.MediaDuration))
                line5 = $"⏱️ Duration: {metadata.MediaDuration}";
            MetadataLine5.Text = line5;
        }
        
        private string FormatFileSize(long bytes)
        {
            if (bytes < 1024) return $"{bytes} B";
            if (bytes < 1024 * 1024) return $"{bytes / 1024.0:F1} KB";
            if (bytes < 1024 * 1024 * 1024) return $"{bytes / (1024.0 * 1024.0):F1} MB";
            return $"{bytes / (1024.0 * 1024.0 * 1024.0):F1} GB";
        }
        
        private string GetFileTypeDescription(string extension)
        {
            return extension.ToLower() switch
            {
                ".pdf" => "PDF Document",
                ".doc" or ".docx" => "Word Document", 
                ".xls" or ".xlsx" => "Excel Spreadsheet",
                ".ppt" or ".pptx" => "PowerPoint",
                ".jpg" or ".jpeg" => "JPEG Image",
                ".png" => "PNG Image",
                ".gif" => "GIF Image",
                ".mp3" => "MP3 Audio",
                ".wav" => "WAV Audio",
                ".mp4" => "MP4 Video",
                ".avi" => "AVI Video",
                ".zip" => "ZIP Archive",
                ".exe" => "Executable",
                ".txt" => "Text File",
                _ => $"{extension.ToUpper()} File"
            };
        }        private void ScanWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {            // Stop scanning animation
            _scanningStoryboard?.Stop();
            _iconCycleTimer?.Stop();
            ScanAnimation.Visibility = Visibility.Collapsed;
              // Reset scan button to normal state
            ScanButton.IsEnabled = true;
            ScanButton.Content = "Scan Files and Export Data to Excel";
            ScanButton.Opacity = 1.0;
              // Reset tooltip back to scan mode
            var tooltip = ScanButton.ToolTip as System.Windows.Controls.ToolTip;
            if (tooltip != null)
            {
                tooltip.Content = "🚀 Start scanning selected folder and export metadata to Excel file";
            }
            
            // Re-enable Select Folder and Select Output File buttons after scanning
            SelectFolderButton.IsEnabled = true;
            BrowseOutputButton.IsEnabled = true;
            
            ProgressBar.Value = 0;
            
            if (e.Error != null)
            {
                WpfMessageBox.Show($"Error during scan: {e.Error.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusTextBlock.Text = "❌ Scan failed";
                StatusBarText.Text = "Scan failed - Please try again";
            }
            else if (e.Cancelled)
            {
                StatusTextBlock.Text = "⏹️ Scan cancelled";
                StatusBarText.Text = "Scan was cancelled by user";
            }
            else if (e.Result is List<FileMetadata> files)
            {
                foreach (var file in files)
                {
                    _fileMetadataList.Add(file);
                }                StatusTextBlock.Text = "✅ Scan completed";
                StatusBarText.Text = $"Scan completed successfully! Found {files.Count} files in {_selectedFolderPath}";
                // Automatically export and open Excel
                if (_fileMetadataList.Count > 0)
                {
                    string outputPath = OutputFileTextBox.Text;
                    try
                    {
                        // Inline ExportToExcelWithFileInUseCheck logic
                        ExportToExcel(outputPath);
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true });
                        StatusTextBlock.Text = "✅ Exported and opened Excel";
                        StatusBarText.Text = "Exported and opened Excel file";
                    }            catch (IOException ioex)
            {
                WpfMessageBox.Show($"❌ Cannot save Excel file because it's currently open!\n\n" +
                                 $"📁 File: {System.IO.Path.GetFileName(outputPath)}\n\n" +
                                 $"💡 Solution: Please close the Excel file and try scanning again.\n\n" +
                                 $"The file might be open in:\n" +
                                 $"• Microsoft Excel\n" +
                                 $"• Another spreadsheet application\n" +
                                 $"• File preview in Windows Explorer\n\n" +
                                 $"Technical details: {ioex.Message}", 
                                 "Excel File Is Open", MessageBoxButton.OK, MessageBoxImage.Warning);
                StatusTextBlock.Text = "❌ Excel file is open";
                StatusBarText.Text = "Excel file is open in another program - Please close it first";
            }
                    catch (Exception ex)
                    {
                        WpfMessageBox.Show($"Error during export: {ex.Message}", "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        StatusTextBlock.Text = "❌ Export failed";
                        StatusBarText.Text = "Export failed - Please try again";
                    }
                }
            }
            else if (e.Result is Exception ex)
            {
                WpfMessageBox.Show($"Error during scan: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusTextBlock.Text = "❌ Scan failed";
                StatusBarText.Text = "Scan failed - Please try again";
            }
        }private void BrowseOutputButton_Click(object sender, RoutedEventArgs e)
        {
            var saveDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                FileName = System.IO.Path.GetFileName(OutputFileTextBox.Text)
            };            if (saveDialog.ShowDialog() == true)
            {
                OutputFileTextBox.Text = saveDialog.FileName;
                // Update file monitoring when output path changes
                UpdateFileMonitoring();
                CheckFileStatus();
            }}        private void ExportToExcel(string filePath)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("File Metadata");

            // Check if hyperlinks should be included
            bool includeHyperlinks = IncludeHyperlinksCheckBox.IsChecked == true;

            // Define core system columns based on hyperlinks setting
            string[] coreHeaders;            if (includeHyperlinks)
            {
                coreHeaders = new[]
                {
                    // Basic Information
                    "File Name", "File Link", "Open Folder", "Full Path", "Directory", "Extension", "File Name (No Ext)",
                    
                    // Size Information  
                    "Size (Bytes)", "Size (KB)", "Size (MB)", "Size (GB)",
                    
                    // Date/Time Information
                    "Created Date", "Modified Date", "Accessed Date", 
                    "Created Date UTC", "Modified Date UTC", "Accessed Date UTC",
                    
                    // File Attributes
                    "Is Read Only", "Is Hidden", "Is System", "Is Archive", 
                    "Is Compressed", "Is Encrypted", "Is Temporary", "All Attributes",
                      // Security & Ownership
                    "Owner", "Access Rights", "Is Executable",
                    
                    // File Hashes
                    "MD5 Hash", "SHA1 Hash", "SHA256 Hash",
                    
                    // File Type & Content
                    "MIME Type", "Registry File Type", "Is Archive File",
                    "Archive File Count", "Archive Uncompressed Size",
                    
                    // System Information
                    "Hard Link Count", "Volume Serial Number", "Drive Type", "File System Type",
                    "Alternate Data Streams"
                };
            }
            else
            {
                coreHeaders = new[]
                {
                    // Basic Information
                    "File Name", "Full Path", "Directory", "Extension", "File Name (No Ext)",
                    
                    // Size Information  
                    "Size (Bytes)", "Size (KB)", "Size (MB)", "Size (GB)",
                    
                    // Date/Time Information
                    "Created Date", "Modified Date", "Accessed Date", 
                    "Created Date UTC", "Modified Date UTC", "Accessed Date UTC",
                      // File Attributes
                    "Is Read Only", "Is Hidden", "Is System", "Is Archive", 
                    "Is Compressed", "Is Encrypted", "Is Temporary", "All Attributes",
                      // Security & Ownership
                    "Owner", "Access Rights", "Is Executable",
                    
                    // File Hashes
                    "MD5 Hash", "SHA1 Hash", "SHA256 Hash",
                    
                    // File Type & Content
                    "MIME Type", "Registry File Type", "Is Archive File",
                    "Archive File Count", "Archive Uncompressed Size",
                    
                    // System Information
                    "Hard Link Count", "Volume Serial Number", "Drive Type", "File System Type",
                    "Alternate Data Streams"
                };
            }// Discover all unique dynamic properties across all files
            var allDynamicProperties = new HashSet<string>();
            foreach (var file in _fileMetadataList)
            {
                foreach (var property in file.DynamicProperties.Keys)
                {
                    allDynamicProperties.Add(property);
                }
            }

            // Create a HashSet of core header names for fast lookup (case-insensitive)
            var coreHeadersSet = new HashSet<string>(coreHeaders, StringComparer.OrdinalIgnoreCase);

            // Filter out dynamic properties that conflict with core headers
            var filteredDynamicProperties = allDynamicProperties
                .Where(prop => !coreHeadersSet.Contains(prop))
                .ToList();

            // Sort dynamic properties alphabetically for consistent column ordering
            var sortedDynamicProperties = filteredDynamicProperties.OrderBy(p => p).ToList();

            // Combine core headers with dynamic property headers (no duplicates)
            var allHeaders = coreHeaders.Concat(sortedDynamicProperties).ToArray();

            // Add headers to the worksheet
            for (int i = 0; i < allHeaders.Length; i++)
            {
                worksheet.Cells[1, i + 1].Value = allHeaders[i];
            }

            // Add data to the worksheet
            for (int row = 0; row < _fileMetadataList.Count; row++)
            {
                var file = _fileMetadataList[row];
                int excelRow = row + 2; // Excel is 1-based, plus header row
                int col = 1;                // Core system columns
                worksheet.Cells[excelRow, col++].Value = file.Name;
                  if (includeHyperlinks)
                {
                    // Add clickable hyperlink to the file
                    var fileLinkCell = worksheet.Cells[excelRow, col++];
                    fileLinkCell.Hyperlink = new Uri(file.FullPath);
                    fileLinkCell.Value = "📁 Open File";
                    fileLinkCell.Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(102, 0, 0)); // Maroon #660000
                    fileLinkCell.Style.Font.UnderLine = true;
                    
                    // Add clickable hyperlink to open the containing folder
                    var folderLinkCell = worksheet.Cells[excelRow, col++];
                    folderLinkCell.Hyperlink = new Uri(file.DirectoryName);
                    folderLinkCell.Value = "📂 Open Folder";
                    folderLinkCell.Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(128, 0, 0)); // Darker maroon #800000
                    folderLinkCell.Style.Font.UnderLine = true;
                }
                
                worksheet.Cells[excelRow, col++].Value = file.FullPath;
                worksheet.Cells[excelRow, col++].Value = file.DirectoryName;
                worksheet.Cells[excelRow, col++].Value = file.Extension;
                worksheet.Cells[excelRow, col++].Value = file.FileNameWithoutExtension;
                
                worksheet.Cells[excelRow, col++].Value = file.SizeBytes;
                worksheet.Cells[excelRow, col++].Value = file.SizeKB;
                worksheet.Cells[excelRow, col++].Value = file.SizeMB;
                worksheet.Cells[excelRow, col++].Value = file.SizeGB;
                
                worksheet.Cells[excelRow, col++].Value = file.CreatedDate;
                worksheet.Cells[excelRow, col++].Value = file.ModifiedDate;
                worksheet.Cells[excelRow, col++].Value = file.AccessedDate;
                worksheet.Cells[excelRow, col++].Value = file.CreatedDateUTC;
                worksheet.Cells[excelRow, col++].Value = file.ModifiedDateUTC;
                worksheet.Cells[excelRow, col++].Value = file.AccessedDateUTC;
                
                worksheet.Cells[excelRow, col++].Value = file.IsReadOnly ? "Yes" : "No";
                worksheet.Cells[excelRow, col++].Value = file.IsHidden ? "Yes" : "No";
                worksheet.Cells[excelRow, col++].Value = file.IsSystem ? "Yes" : "No";
                worksheet.Cells[excelRow, col++].Value = file.IsArchive ? "Yes" : "No";
                worksheet.Cells[excelRow, col++].Value = file.IsCompressed ? "Yes" : "No";
                worksheet.Cells[excelRow, col++].Value = file.IsEncrypted ? "Yes" : "No";
                worksheet.Cells[excelRow, col++].Value = file.IsTemporary ? "Yes" : "No";
                worksheet.Cells[excelRow, col++].Value = file.Attributes;
                  worksheet.Cells[excelRow, col++].Value = file.Owner;
                worksheet.Cells[excelRow, col++].Value = file.AccessRights;                worksheet.Cells[excelRow, col++].Value = file.IsExecutable ? "Yes" : "No";
                
                worksheet.Cells[excelRow, col++].Value = file.MD5Hash;
                worksheet.Cells[excelRow, col++].Value = file.SHA1Hash;
                worksheet.Cells[excelRow, col++].Value = file.SHA256Hash;
                
                worksheet.Cells[excelRow, col++].Value = file.MimeType;
                worksheet.Cells[excelRow, col++].Value = file.RegistryFileType;
                worksheet.Cells[excelRow, col++].Value = file.IsArchiveFile ? "Yes" : "No";
                worksheet.Cells[excelRow, col++].Value = file.ArchiveFileCount;
                worksheet.Cells[excelRow, col++].Value = file.ArchiveUncompressedSize;
                
                worksheet.Cells[excelRow, col++].Value = file.HardLinkCount;
                worksheet.Cells[excelRow, col++].Value = file.VolumeSerialNumber;
                worksheet.Cells[excelRow, col++].Value = file.DriveType;
                worksheet.Cells[excelRow, col++].Value = file.FileSystemType;
                worksheet.Cells[excelRow, col++].Value = file.AlternateDataStreams;                // Dynamic properties columns
                foreach (var propertyName in sortedDynamicProperties)
                {
                    var value = file.DynamicProperties.TryGetValue(propertyName, out var propValue) ? propValue : "";
                    worksheet.Cells[excelRow, col++].Value = value;
                }
            }

            // Create a formal Excel table with filter buttons
            var dataRange = worksheet.Cells[1, 1, _fileMetadataList.Count + 1, allHeaders.Length];
            var table = worksheet.Tables.Add(dataRange, "FileMetadataTable");
            
            // Set table style - professional blue theme
            table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
            table.ShowFilter = true;
            table.ShowHeader = true;
            table.ShowTotal = false;            // Format the table headers
            var headerRange = worksheet.Cells[1, 1, 1, allHeaders.Length];
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Font.Color.SetColor(System.Drawing.Color.White);
            headerRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(102, 0, 0)); // Maroon #660000
            headerRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;// Format specific column types
            if (_fileMetadataList.Count > 0)
            {
                if (includeHyperlinks)
                {
                    // Date columns (12-14) - with hyperlink columns
                    var createdDateCol = worksheet.Cells[2, 12, _fileMetadataList.Count + 1, 12];
                    createdDateCol.Style.Numberformat.Format = "mm/dd/yyyy hh:mm AM/PM";
                    
                    var modifiedDateCol = worksheet.Cells[2, 13, _fileMetadataList.Count + 1, 13];
                    modifiedDateCol.Style.Numberformat.Format = "mm/dd/yyyy hh:mm AM/PM";
                    
                    var accessedDateCol = worksheet.Cells[2, 14, _fileMetadataList.Count + 1, 14];
                    accessedDateCol.Style.Numberformat.Format = "mm/dd/yyyy hh:mm AM/PM";

                    // Size columns (8-11) - with hyperlink columns
                    var sizesBytesCol = worksheet.Cells[2, 8, _fileMetadataList.Count + 1, 8];
                    sizesBytesCol.Style.Numberformat.Format = "#,##0";
                    
                    var sizesKBCol = worksheet.Cells[2, 9, _fileMetadataList.Count + 1, 9];
                    sizesKBCol.Style.Numberformat.Format = "#,##0.00";
                    
                    var sizesMBCol = worksheet.Cells[2, 10, _fileMetadataList.Count + 1, 10];
                    sizesMBCol.Style.Numberformat.Format = "#,##0.00";
                    
                    var sizesGBCol = worksheet.Cells[2, 11, _fileMetadataList.Count + 1, 11];
                    sizesGBCol.Style.Numberformat.Format = "#,##0.0000";
                }
                else
                {
                    // Date columns (10-12) - without hyperlink columns
                    var createdDateCol = worksheet.Cells[2, 10, _fileMetadataList.Count + 1, 10];
                    createdDateCol.Style.Numberformat.Format = "mm/dd/yyyy hh:mm AM/PM";
                    
                    var modifiedDateCol = worksheet.Cells[2, 11, _fileMetadataList.Count + 1, 11];
                    modifiedDateCol.Style.Numberformat.Format = "mm/dd/yyyy hh:mm AM/PM";
                    
                    var accessedDateCol = worksheet.Cells[2, 12, _fileMetadataList.Count + 1, 12];
                    accessedDateCol.Style.Numberformat.Format = "mm/dd/yyyy hh:mm AM/PM";

                    // Size columns (6-9) - without hyperlink columns
                    var sizesBytesCol = worksheet.Cells[2, 6, _fileMetadataList.Count + 1, 6];
                    sizesBytesCol.Style.Numberformat.Format = "#,##0";
                    
                    var sizesKBCol = worksheet.Cells[2, 7, _fileMetadataList.Count + 1, 7];
                    sizesKBCol.Style.Numberformat.Format = "#,##0.00";
                    
                    var sizesMBCol = worksheet.Cells[2, 8, _fileMetadataList.Count + 1, 8];
                    sizesMBCol.Style.Numberformat.Format = "#,##0.00";
                    
                    var sizesGBCol = worksheet.Cells[2, 9, _fileMetadataList.Count + 1, 9];
                    sizesGBCol.Style.Numberformat.Format = "#,##0.0000";
                }
            }            // Auto-fit columns and set minimum widths
            worksheet.Cells.AutoFitColumns();
            // Ensure columns fit header text (including filter dropdown)
            for (int i = 1; i <= allHeaders.Length; i++)
            {
                var headerText = worksheet.Cells[1, i].Value?.ToString() ?? string.Empty;
                double headerWidthEstimate = headerText.Length + 8; // add extra padding for filter button
                if (worksheet.Column(i).Width < headerWidthEstimate)
                    worksheet.Column(i).Width = headerWidthEstimate;
            }

            // Set minimum column widths for readability
            for (int i = 1; i <= allHeaders.Length; i++)
            {
                if (worksheet.Column(i).Width < 10) worksheet.Column(i).Width = 10;
                if (worksheet.Column(i).Width > 50) worksheet.Column(i).Width = 50; // Cap very wide columns
            }            // Set specific widths for key columns
            if (includeHyperlinks)
            {
                worksheet.Column(1).Width = Math.Max(worksheet.Column(1).Width, 20); // File Name
                worksheet.Column(2).Width = Math.Max(worksheet.Column(2).Width, 15); // File Link
                worksheet.Column(3).Width = Math.Max(worksheet.Column(3).Width, 15); // Open Folder
                worksheet.Column(4).Width = Math.Max(worksheet.Column(4).Width, 50); // Full Path
                worksheet.Column(5).Width = Math.Max(worksheet.Column(5).Width, 30); // Directory
            }
            else
            {
                worksheet.Column(1).Width = Math.Max(worksheet.Column(1).Width, 20); // File Name
                worksheet.Column(2).Width = Math.Max(worksheet.Column(2).Width, 50); // Full Path
                worksheet.Column(3).Width = Math.Max(worksheet.Column(3).Width, 30); // Directory
            }

            // Add borders to the entire table
            dataRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            dataRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            // Freeze the header row
            worksheet.View.FreezePanes(2, 1);

            // Add summary information on a separate sheet
            var summaryWorksheet = package.Workbook.Worksheets.Add("Summary");            summaryWorksheet.Cells[1, 1].Value = "File Metadata Scan Summary";
            summaryWorksheet.Cells[1, 1].Style.Font.Bold = true;
            summaryWorksheet.Cells[1, 1].Style.Font.Size = 18;
            summaryWorksheet.Cells[1, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(102, 0, 0)); // Maroon #660000

            // Summary data
            summaryWorksheet.Cells[3, 1].Value = "Scanned Folder:";
            summaryWorksheet.Cells[3, 2].Value = _selectedFolderPath;
            summaryWorksheet.Cells[4, 1].Value = "Total Files Found:";
            summaryWorksheet.Cells[4, 2].Value = _fileMetadataList.Count;
            summaryWorksheet.Cells[5, 1].Value = "Scan Date:";
            summaryWorksheet.Cells[5, 2].Value = DateTime.Now;
            summaryWorksheet.Cells[6, 1].Value = "Total Size (MB):";
            summaryWorksheet.Cells[6, 2].Value = Math.Round(_fileMetadataList.Sum(f => f.SizeMB), 2);            summaryWorksheet.Cells[7, 1].Value = "Dynamic Properties Found:";
            summaryWorksheet.Cells[7, 2].Value = sortedDynamicProperties.Count;
            summaryWorksheet.Cells[8, 1].Value = "Total Columns:";
            summaryWorksheet.Cells[8, 2].Value = allHeaders.Length;

            // Format summary labels
            summaryWorksheet.Cells[3, 1, 8, 1].Style.Font.Bold = true;
            summaryWorksheet.Cells[3, 1, 8, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            summaryWorksheet.Cells[3, 1, 8, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(217, 225, 242));

            // Format summary date
            summaryWorksheet.Cells[5, 2].Style.Numberformat.Format = "mm/dd/yyyy hh:mm AM/PM";
            summaryWorksheet.Cells[6, 2].Style.Numberformat.Format = "#,##0.00";

            summaryWorksheet.Cells.AutoFitColumns();

            // Set the File Metadata sheet as the active sheet
            package.Workbook.Worksheets["File Metadata"].Select();

            package.SaveAs(new FileInfo(filePath));
        }
        
        // Comprehensive metadata extraction method
        private FileMetadata ExtractComprehensiveMetadata(string filePath, bool diagnosticMode)
        {
            var metadata = new FileMetadata();
            
            try
            {
                // Basic file information
                var fileName = Path.GetFileName(filePath);
                var directory = Path.GetDirectoryName(filePath) ?? "";
                var extension = Path.GetExtension(filePath);
                
                metadata.Name = fileName;
                metadata.FullPath = filePath;
                metadata.DirectoryName = directory;
                metadata.Extension = extension;
                metadata.FileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
                
                if (diagnosticMode)
                {
                    // Use minimal file access
                    var attributes = File.GetAttributes(filePath);
                    var creationTime = File.GetCreationTime(filePath);
                    var lastWriteTime = File.GetLastWriteTime(filePath);
                    var lastAccessTime = File.GetLastAccessTime(filePath);
                    
                    metadata.CreatedDate = creationTime;
                    metadata.ModifiedDate = lastWriteTime;
                    metadata.AccessedDate = lastAccessTime;
                    metadata.CreatedDateUTC = creationTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss UTC");
                    metadata.ModifiedDateUTC = lastWriteTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss UTC");
                    metadata.AccessedDateUTC = lastAccessTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss UTC");
                    metadata.Attributes = attributes.ToString();
                    
                    // File attributes
                    metadata.IsReadOnly = (attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly;
                    metadata.IsHidden = (attributes & FileAttributes.Hidden) == FileAttributes.Hidden;
                    metadata.IsSystem = (attributes & FileAttributes.System) == FileAttributes.System;
                    metadata.IsArchive = (attributes & FileAttributes.Archive) == FileAttributes.Archive;
                    metadata.IsCompressed = (attributes & FileAttributes.Compressed) == FileAttributes.Compressed;
                    metadata.IsEncrypted = (attributes & FileAttributes.Encrypted) == FileAttributes.Encrypted;
                    metadata.IsTemporary = (attributes & FileAttributes.Temporary) == FileAttributes.Temporary;
                    
                    // Get file size
                    try
                    {
                        var fileInfo = new FileInfo(filePath);
                        var length = fileInfo.Length;
                        metadata.SizeBytes = length;
                        metadata.SizeKB = Math.Round(length / 1024.0, 2);
                        metadata.SizeMB = Math.Round(length / (1024.0 * 1024.0), 2);
                        metadata.SizeGB = Math.Round(length / (1024.0 * 1024.0 * 1024.0), 4);
                    }
                    catch { /* Size will remain 0 */ }
                }
                else
                {
                    // Use FileInfo for comprehensive data
                    var fileInfo = new FileInfo(filePath);
                    
                    metadata.SizeBytes = fileInfo.Length;
                    metadata.SizeKB = Math.Round(fileInfo.Length / 1024.0, 2);
                    metadata.SizeMB = Math.Round(fileInfo.Length / (1024.0 * 1024.0), 2);
                    metadata.SizeGB = Math.Round(fileInfo.Length / (1024.0 * 1024.0 * 1024.0), 4);
                    
                    metadata.CreatedDate = fileInfo.CreationTime;
                    metadata.ModifiedDate = fileInfo.LastWriteTime;
                    metadata.AccessedDate = fileInfo.LastAccessTime;
                    metadata.CreatedDateUTC = fileInfo.CreationTimeUtc.ToString("yyyy-MM-dd HH:mm:ss UTC");
                    metadata.ModifiedDateUTC = fileInfo.LastWriteTimeUtc.ToString("yyyy-MM-dd HH:mm:ss UTC");
                    metadata.AccessedDateUTC = fileInfo.LastAccessTimeUtc.ToString("yyyy-MM-dd HH:mm:ss UTC");
                    
                    var attributes = fileInfo.Attributes;
                    metadata.Attributes = attributes.ToString();
                    metadata.IsReadOnly = fileInfo.IsReadOnly;
                    metadata.IsHidden = (attributes & FileAttributes.Hidden) == FileAttributes.Hidden;
                    metadata.IsSystem = (attributes & FileAttributes.System) == FileAttributes.System;
                    metadata.IsArchive = (attributes & FileAttributes.Archive) == FileAttributes.Archive;
                    metadata.IsCompressed = (attributes & FileAttributes.Compressed) == FileAttributes.Compressed;
                    metadata.IsEncrypted = (attributes & FileAttributes.Encrypted) == FileAttributes.Encrypted;
                    metadata.IsTemporary = (attributes & FileAttributes.Temporary) == FileAttributes.Temporary;
                }
                
                // Drive and file system information
                try
                {
                    var driveInfo = new DriveInfo(Path.GetPathRoot(filePath) ?? "");
                    metadata.DriveType = driveInfo.DriveType.ToString();
                    metadata.FileSystemType = driveInfo.DriveFormat;
                    metadata.VolumeSerialNumber = driveInfo.Name;
                }                catch { /* Drive info not available */ }

                // Hash computation for files under 100MB (valuable for duplicate detection and file integrity)
                if (metadata.SizeBytes < 100 * 1024 * 1024) // 100MB limit
                {
                    try
                    {
                        using var stream = File.OpenRead(filePath);
                        
                        // Calculate MD5
                        using (var md5 = MD5.Create())
                        {
                            stream.Position = 0;
                            var hashBytes = md5.ComputeHash(stream);
                            metadata.MD5Hash = Convert.ToHexString(hashBytes);
                        }
                        
                        // Calculate SHA1
                        using (var sha1 = SHA1.Create())
                        {
                            stream.Position = 0;
                            var hashBytes = sha1.ComputeHash(stream);
                            metadata.SHA1Hash = Convert.ToHexString(hashBytes);
                        }
                        
                        // Calculate SHA256
                        using (var sha256 = SHA256.Create())
                        {
                            stream.Position = 0;
                            var hashBytes = sha256.ComputeHash(stream);
                            metadata.SHA256Hash = Convert.ToHexString(hashBytes);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Hash computation failed, but continue with other metadata
                        System.Diagnostics.Debug.WriteLine($"Hash computation failed for {filePath}: {ex.Message}");
                        metadata.MD5Hash = "Error computing hash";
                        metadata.SHA1Hash = "Error computing hash";
                        metadata.SHA256Hash = "Error computing hash";
                    }
                }
                else
                {
                    // File too large for hash computation
                    metadata.MD5Hash = "File too large (>100MB)";
                    metadata.SHA1Hash = "File too large (>100MB)";
                    metadata.SHA256Hash = "File too large (>100MB)";
                }

                // Security information
                try
                {
                    var fileInfo = new FileInfo(filePath);
                    var security = fileInfo.GetAccessControl();
                    metadata.Owner = security.GetOwner(typeof(NTAccount))?.ToString() ?? "";
                    
                    var accessRules = security.GetAccessRules(true, true, typeof(NTAccount));
                    var rights = new List<string>();
                    foreach (FileSystemAccessRule rule in accessRules)
                    {
                        rights.Add($"{rule.IdentityReference}: {rule.FileSystemRights}");
                    }
                    metadata.AccessRights = string.Join("; ", rights.Take(5)); // Limit to first 5 rules
                }
                catch 
                { 
                    // Security info not available in .NET Core - use alternative approach
                    try
                    {
                        // Get basic owner info using WMI or alternative methods
                        metadata.Owner = "Not available in .NET Core";
                        metadata.AccessRights = "Requires .NET Framework";
                    }
                    catch { /* Security info not available */ }
                }
                
                // Executable detection
                var executableExtensions = new[] { ".exe", ".dll", ".com", ".bat", ".cmd", ".msi" };
                metadata.IsExecutable = executableExtensions.Contains(extension.ToLower());
                
                // Version information for executables
                if (metadata.IsExecutable)
                {
                    try
                    {
                        var versionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(filePath);
                        metadata.FileDescription = versionInfo.FileDescription ?? "";
                        metadata.FileVersion = versionInfo.FileVersion ?? "";
                        metadata.ProductName = versionInfo.ProductName ?? "";
                        metadata.ProductVersion = versionInfo.ProductVersion ?? "";
                        metadata.CompanyName = versionInfo.CompanyName ?? "";
                        metadata.Copyright = versionInfo.LegalCopyright ?? "";
                        metadata.InternalName = versionInfo.InternalName ?? "";
                        metadata.OriginalFileName = versionInfo.OriginalFilename ?? "";
                        metadata.Comments = versionInfo.Comments ?? "";
                        metadata.LegalTrademarks = versionInfo.LegalTrademarks ?? "";
                    }
                    catch { /* Version info not available */ }
                }
                  // MIME type detection
                metadata.MimeType = GetMimeType(extension);
                
                // Extract comprehensive Windows Shell properties (what File Properties shows)
                ExtractWindowsShellProperties(filePath, metadata);
                
                // Archive file detection
                var archiveExtensions = new[] { ".zip", ".rar", ".7z", ".tar", ".gz", ".bz2" };
                metadata.IsArchiveFile = archiveExtensions.Contains(extension.ToLower());
                
                // Registry file type information
                try
                {
                    using var key = Registry.ClassesRoot.OpenSubKey(extension);
                    if (key != null)
                    {
                        metadata.RegistryFileType = key.GetValue("")?.ToString() ?? "";
                    }                }
                catch { /* Registry info not available */ }
                
            }
            catch (Exception ex)
            {
                // Log error but continue
                System.Diagnostics.Debug.WriteLine($"Error extracting metadata for {filePath}: {ex.Message}");
            }
            
            return metadata;
        }
        
        private string GetMimeType(string extension)
        {
            return extension.ToLower() switch
            {
                ".txt" => "text/plain",
                ".html" => "text/html",
                ".css" => "text/css",
                ".js" => "application/javascript",
                ".json" => "application/json",
                ".xml" => "application/xml",
                ".pdf" => "application/pdf",
                ".doc" => "application/msword",
                ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".xls" => "application/vnd.ms-excel",
                ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ".ppt" => "application/vnd.ms-powerpoint",
                ".pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                ".jpg" or ".jpeg" => "image/jpeg",
                ".png" => "image/png",
                ".gif" => "image/gif",
                ".bmp" => "image/bmp",
                ".ico" => "image/x-icon",
                ".svg" => "image/svg+xml",
                ".mp3" => "audio/mpeg",
                ".wav" => "audio/wav",
                ".mp4" => "video/mp4",
                ".avi" => "video/x-msvideo",
                ".mov" => "video/quicktime",
                ".zip" => "application/zip",
                ".rar" => "application/x-rar-compressed",                ".exe" => "application/x-msdownload",
                ".dll" => "application/x-msdownload",
                _ => "application/octet-stream"
            };
        }
          // Enhanced method to extract ALL Windows Shell properties (comprehensive file properties)
        private void ExtractWindowsShellProperties(string filePath, FileMetadata metadata)
        {
            try
            {
                // Use Shell32 to get ALL extended properties
                dynamic shell = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));
                var folder = shell.NameSpace(Path.GetDirectoryName(filePath));
                var folderItem = folder.ParseName(Path.GetFileName(filePath));
                
                if (folderItem != null)
                {
                    // Store all properties in a dictionary for complete capture
                    var allProperties = new Dictionary<string, string>();                    // Extract ALL available shell properties (Windows supports up to ~300+ properties)
                    for (int i = 0; i < 500; i++) // Increased range to capture everything
                    {
                        try
                        {
                            string value = folder.GetDetailsOf(folderItem, i)?.ToString();
                            if (!string.IsNullOrWhiteSpace(value))
                            {
                                string propertyName = folder.GetDetailsOf(null, i)?.ToString() ?? $"Property_{i}";
                                
                                // Clean up property name
                                propertyName = propertyName.Trim();
                                if (!string.IsNullOrEmpty(propertyName) && propertyName != "Property_" + i)
                                {
                                    // Map to our metadata fields and store in DynamicProperties for individual columns
                                    MapPropertyToMetadata(propertyName, value, metadata);
                                }
                            }
                        }
                        catch
                        {
                            // Skip properties that can't be accessed
                        }
                    }
                }
                
                // Clean up COM objects
                if (folderItem != null) Marshal.ReleaseComObject(folderItem);
                if (folder != null) Marshal.ReleaseComObject(folder);
                if (shell != null) Marshal.ReleaseComObject(shell);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ExtractWindowsShellProperties failed for {filePath}: {ex.Message}");
            }
        }        // Comprehensive property mapping method
        private void MapPropertyToMetadata(string propertyName, string value, FileMetadata metadata)
        {
            var lowerName = propertyName.ToLower();
              // Store ALL properties in DynamicProperties (we'll create individual columns for everything)
            // No exclusions - we want every property to have its own column
            if (!string.IsNullOrWhiteSpace(value) && value.Length < 500)
            {
                // Clean the property name to avoid issues
                var cleanPropertyName = propertyName.Trim();
                if (!string.IsNullOrEmpty(cleanPropertyName))
                {
                    // Just use the key - if it already exists, overwrite it
                    metadata.DynamicProperties[cleanPropertyName] = value;
                }
            }
            switch (lowerName)
            {
                // Document Properties (also stored in DynamicProperties)
                case "author":
                case "authors":
                    metadata.DocumentAuthor = SetIfEmpty(metadata.DocumentAuthor, value);
                    break;
                case "title":
                    metadata.DocumentTitle = SetIfEmpty(metadata.DocumentTitle, value);
                    break;
                case "subject":
                    metadata.DocumentSubject = SetIfEmpty(metadata.DocumentSubject, value);
                    break;
                case "keywords":
                case "tags":
                    metadata.DocumentKeywords = SetIfEmpty(metadata.DocumentKeywords, value);
                    break;
                case "comments":
                    metadata.Comments = SetIfEmpty(metadata.Comments, value);
                    break;
                case "company":
                    metadata.CompanyName = SetIfEmpty(metadata.CompanyName, value);
                    break;
                case "copyright":
                    metadata.Copyright = SetIfEmpty(metadata.Copyright, value);
                    break;
                case "last saved by":
                    metadata.LastSavedBy = SetIfEmpty(metadata.LastSavedBy, value);
                    break;
                case "application name":
                case "program name":
                    metadata.ApplicationName = SetIfEmpty(metadata.ApplicationName, value);
                    break;
                case "pages":
                    if (metadata.PageCount == 0 && int.TryParse(value, out int pages))
                        metadata.PageCount = pages;
                    break;
                case "word count":
                    if (metadata.WordCount == 0 && int.TryParse(value, out int words))
                        metadata.WordCount = words;
                    break;
                case "character count":
                    if (metadata.CharacterCount == 0 && int.TryParse(value, out int chars))
                        metadata.CharacterCount = chars;
                    break;
                case "content type":
                case "perceived type":
                case "kind":
                    metadata.ContentType = SetIfEmpty(metadata.ContentType, value);
                    break;

                // Image/Photo Properties
                case "camera model":
                    metadata.CameraModel = SetIfEmpty(metadata.CameraModel, value);
                    break;
                case "camera make":
                case "camera manufacturer":
                case "camera maker":
                    metadata.CameraMake = SetIfEmpty(metadata.CameraMake, value);
                    break;
                case "date picture taken":
                case "date taken":
                    metadata.DateTaken = SetIfEmpty(metadata.DateTaken, value);
                    break;
                case "dimensions":
                    if (string.IsNullOrEmpty(metadata.ImageWidth))
                    {
                        var parts = value.Split(new char[] { 'x', '×', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length >= 2)
                        {
                            metadata.ImageWidth = parts[0].Trim();
                            metadata.ImageHeight = parts[1].Trim();
                        }
                    }
                    break;
                case "width":
                case "image width":
                    metadata.ImageWidth = SetIfEmpty(metadata.ImageWidth, value);
                    break;
                case "height":
                case "image height":
                    metadata.ImageHeight = SetIfEmpty(metadata.ImageHeight, value);
                    break;
                case "horizontal resolution":
                    metadata.HorizontalResolution = SetIfEmpty(metadata.HorizontalResolution, value);
                    break;
                case "vertical resolution":
                    metadata.VerticalResolution = SetIfEmpty(metadata.VerticalResolution, value);
                    break;
                case "bit depth":
                case "color depth":
                    metadata.BitDepth = SetIfEmpty(metadata.BitDepth, value);
                    break;
                case "f-stop":
                case "f-number":
                    metadata.FNumber = SetIfEmpty(metadata.FNumber, value);
                    break;
                case "exposure time":
                    metadata.ExposureTime = SetIfEmpty(metadata.ExposureTime, value);
                    break;
                case "iso speed":
                case "iso":
                    metadata.ISO = SetIfEmpty(metadata.ISO, value);
                    break;
                case "focal length":
                    metadata.FocalLength = SetIfEmpty(metadata.FocalLength, value);
                    break;
                case "flash":
                case "flash mode":
                    metadata.Flash = SetIfEmpty(metadata.Flash, value);
                    break;
                case "orientation":
                    metadata.Orientation = SetIfEmpty(metadata.Orientation, value);
                    break;
                case "gps latitude":
                case "latitude":
                    metadata.GPS_Latitude = SetIfEmpty(metadata.GPS_Latitude, value);
                    break;
                case "gps longitude":
                case "longitude":
                    metadata.GPS_Longitude = SetIfEmpty(metadata.GPS_Longitude, value);
                    break;                case "rating":
                    // Skip rating if it's "Unrated" 
                    if (value != "Unrated")
                    {
                        // Properties are now automatically stored in DynamicProperties above
                    }
                    break;

                // Audio/Video Properties
                case "duration":
                case "length":
                    metadata.MediaDuration = SetIfEmpty(metadata.MediaDuration, value);
                    break;
                case "bit rate":
                case "audio bit rate":
                case "bitrate":
                    metadata.AudioBitrate = SetIfEmpty(metadata.AudioBitrate, value);
                    break;
                case "frame rate":
                case "video frame rate":
                    metadata.FrameRate = SetIfEmpty(metadata.FrameRate, value);
                    break;
                case "frame width":
                case "video width":
                    metadata.FrameWidth = SetIfEmpty(metadata.FrameWidth, value);
                    break;
                case "frame height":
                case "video height":
                    metadata.FrameHeight = SetIfEmpty(metadata.FrameHeight, value);
                    break;
                case "video compression":
                case "video codec":
                    metadata.VideoCodec = SetIfEmpty(metadata.VideoCodec, value);
                    break;
                case "audio compression":
                case "audio codec":
                    metadata.AudioCodec = SetIfEmpty(metadata.AudioCodec, value);
                    break;
                case "album":
                case "album title":
                    metadata.Album = SetIfEmpty(metadata.Album, value);
                    break;
                case "artist":
                case "contributing artists":
                    metadata.Artist = SetIfEmpty(metadata.Artist, value);
                    break;
                case "genre":
                    metadata.Genre = SetIfEmpty(metadata.Genre, value);
                    break;
                case "year":
                case "release year":
                    metadata.Year = SetIfEmpty(metadata.Year, value);
                    break;
                case "track":
                case "track number":
                    metadata.Track = SetIfEmpty(metadata.Track, value);
                    break;

                // Video specific
                case "aspect ratio":
                    metadata.AspectRatio = SetIfEmpty(metadata.AspectRatio, value);
                    break;

                // File System Properties  
                case "owner":
                    metadata.Owner = SetIfEmpty(metadata.Owner, value);
                    break;
                case "file version":
                    metadata.FileVersion = SetIfEmpty(metadata.FileVersion, value);
                    break;
                case "product name":
                    metadata.ProductName = SetIfEmpty(metadata.ProductName, value);
                    break;
                case "product version":
                    metadata.ProductVersion = SetIfEmpty(metadata.ProductVersion, value);
                    break;                case "language":
                    metadata.Language = SetIfEmpty(metadata.Language, value);
                    break;                case "file description":
                case "description":
                    metadata.FileDescription = SetIfEmpty(metadata.FileDescription, value);
                    break;
                
                // All other properties are automatically stored in DynamicProperties at the top of this method
                default:
                    // Properties are already stored in DynamicProperties at the beginning of the method
                    break;
            }
        }// Helper method to set value only if current is empty
        private string SetIfEmpty(string currentValue, string newValue)
        {
            return string.IsNullOrEmpty(currentValue) ? newValue : currentValue;
        }

        private void InitializeFileMonitoring()
        {
            // Timer to check file status periodically
            _fileStatusTimer = new DispatcherTimer();
            _fileStatusTimer.Interval = TimeSpan.FromSeconds(2); // Check every 2 seconds
            _fileStatusTimer.Tick += FileStatusTimer_Tick;
            _fileStatusTimer.Start();
        }

        private void OutputFileTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Update file monitoring when output path changes
            UpdateFileMonitoring();
            CheckFileStatus();
        }

        private void UpdateFileMonitoring()
        {
            try
            {
                // Dispose existing watcher
                _fileWatcher?.Dispose();
                
                string outputPath = OutputFileTextBox.Text;
                if (string.IsNullOrWhiteSpace(outputPath))
                    return;
                    
                string directory = Path.GetDirectoryName(outputPath);
                string fileName = Path.GetFileName(outputPath);
                
                if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory))
                {
                    _fileWatcher = new FileSystemWatcher(directory, fileName);
                    _fileWatcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.CreationTime | NotifyFilters.LastAccess;
                    _fileWatcher.Changed += FileWatcher_Changed;
                    _fileWatcher.Created += FileWatcher_Changed;
                    _fileWatcher.Deleted += FileWatcher_Changed;
                    _fileWatcher.EnableRaisingEvents = true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"File monitoring setup failed: {ex.Message}");
            }
        }

        private void FileWatcher_Changed(object sender, FileSystemEventArgs e)
        {
            // Check file status when changes are detected
            Dispatcher.BeginInvoke(() => CheckFileStatus());
        }

        private void FileStatusTimer_Tick(object sender, EventArgs e)
        {
            // Periodic file status check
            CheckFileStatus();
        }

        private void CheckFileStatus()
        {
            string outputPath = OutputFileTextBox.Text;
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                UpdateUIForFileStatus(false, "");
                return;
            }
            
            bool wasInUse = _isExcelFileInUse;
            _isExcelFileInUse = IsFileInUse(outputPath);
            
            // Only update UI if status changed
            if (wasInUse != _isExcelFileInUse)
            {
                UpdateUIForFileStatus(_isExcelFileInUse, outputPath);
            }
        }

        private bool IsFileInUse(string filePath)
        {
            if (!File.Exists(filePath))
                return false;
                
            try
            {
                using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    // If we can open with exclusive access, file is not in use
                }
                return false;
            }
            catch (IOException)
            {
                // File is in use
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                // Treat permission issues as file in use
                return true;
            }
            catch
            {
                // Other errors, assume file is available
                return false;
            }
        }        private void UpdateUIForFileStatus(bool isInUse, string filePath)
        {
            string fileName = string.IsNullOrEmpty(filePath) ? "" : Path.GetFileName(filePath);
            
            Dispatcher.Invoke(() =>
            {
                if (isInUse)
                {
                    // File is in use - disable scanning
                    if (ScanButton != null)
                    {
                        ScanButton.IsEnabled = false;
                        ScanButton.Content = "Excel File Open - Close it to enable scanning";
                        ScanButton.Opacity = 0.6;
                    }
                    
                    if (!string.IsNullOrEmpty(_selectedFolderPath))
                    {
                        StatusBarText.Text = $"📁 Folder selected but Excel file '{fileName}' is open - Close it to enable scanning";
                    }
                    else
                    {
                        StatusBarText.Text = $"⚠️ Excel file '{fileName}' is open - Close it to enable scanning";
                    }
                    StatusTextBlock.Text = "📄 Excel file is open";
                }
                else
                {
                    // File is available - enable scanning if folder is selected
                    if (ScanButton != null && !string.IsNullOrEmpty(_selectedFolderPath) && Directory.Exists(_selectedFolderPath))
                    {
                        ScanButton.IsEnabled = true;
                        ScanButton.Content = "Scan";
                        ScanButton.Opacity = 1.0;
                        ScanButton.Visibility = Visibility.Visible;
                          StatusBarText.Text = $"✅ Ready to scan '{Path.GetFileName(_selectedFolderPath)}' - Excel file available";
                        StatusTextBlock.Text = "✅ Ready to scan";
                    }
                    else if (string.IsNullOrEmpty(_selectedFolderPath))
                    {
                        // No folder selected yet
                        StatusBarText.Text = "📁 Please select a folder using the Browse button";
                        StatusTextBlock.Text = "Ready to scan";
                    }
                }
            });
        }

        protected override void OnClosed(EventArgs e)
        {
            // Clean up resources
            _fileWatcher?.Dispose();
            _fileStatusTimer?.Stop();
            base.OnClosed(e);
        }

        // Event handler for double-clicking the folder path text box to open the folder
        private void FolderPathTextBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (!string.IsNullOrEmpty(_selectedFolderPath) && IODirectory.Exists(_selectedFolderPath))
            {
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(_selectedFolderPath) { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    WpfMessageBox.Show($"Could not open folder: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        // Event handler for double-clicking the output file text box to open the containing folder
        private void OutputFileTextBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            string outputPath = OutputFileTextBox.Text;
            if (!string.IsNullOrEmpty(outputPath))
            {
                try
                {
                    // If the file exists, open it directly; otherwise open the containing folder
                    if (File.Exists(outputPath))
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true });
                    }
                    else
                    {
                        // Open the containing folder
                        string folderPath = Path.GetDirectoryName(outputPath);
                        if (!string.IsNullOrEmpty(folderPath) && IODirectory.Exists(folderPath))
                        {
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(folderPath) { UseShellExecute = true });
                        }
                        else
                        {
                            WpfMessageBox.Show("The output file path does not exist yet.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                    }
                }
                catch (Exception ex)
                {
                    WpfMessageBox.Show($"Could not open path: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }
    }    public class ProgressInfo
    {
        public int FilesProcessed { get; set; }
        public int TotalFiles { get; set; }
        public string CurrentFile { get; set; } = "";
        public int FilesFound { get; set; }
        public FileMetadata CurrentMetadata { get; set; } = null;
    }

    public class FileMetadata
    {
        // Basic File Information
        public string Name { get; set; } = "";
        public string FullPath { get; set; } = "";
        public string DirectoryName { get; set; } = "";
        public string Extension { get; set; } = "";
        public string FileNameWithoutExtension { get; set; } = "";
        
        // Size Information
        public long SizeBytes { get; set; }
        public double SizeKB { get; set; }
        public double SizeMB { get; set; }
        public double SizeGB { get; set; }
        
        // Date/Time Information
        public DateTime CreatedDate { get; set; }
        public DateTime ModifiedDate { get; set; }
        public DateTime AccessedDate { get; set; }
        public string CreatedDateUTC { get; set; } = "";
        public string ModifiedDateUTC { get; set; } = "";
        public string AccessedDateUTC { get; set; } = "";
        
        // File Attributes & Properties
        public bool IsReadOnly { get; set; }
        public bool IsHidden { get; set; }
        public bool IsSystem { get; set; }
        public bool IsArchive { get; set; }
        public bool IsCompressed { get; set; }
        public bool IsEncrypted { get; set; }
        public bool IsTemporary { get; set; }
        public string Attributes { get; set; } = "";
        
        // Security Information
        public string Owner { get; set; } = "";
        public string AccessRights { get; set; } = "";
        public bool IsExecutable { get; set; }
        
        // File Hashes
        public string MD5Hash { get; set; } = "";
        public string SHA1Hash { get; set; } = "";
        public string SHA256Hash { get; set; } = "";
        
        // File Type & Content Information
        public string MimeType { get; set; } = "";
        public string FileDescription { get; set; } = "";
        public string FileVersion { get; set; } = "";
        public string ProductName { get; set; } = "";
        public string ProductVersion { get; set; } = "";
        public string CompanyName { get; set; } = "";
        public string Copyright { get; set; } = "";
        public string InternalName { get; set; } = "";
        public string OriginalFileName { get; set; } = "";
        public string Comments { get; set; } = "";
        public string LegalTrademarks { get; set; } = "";
        
        // Image/Media Specific Metadata (EXIF)
        public string ImageWidth { get; set; } = "";
        public string ImageHeight { get; set; } = "";
        public string ColorDepth { get; set; } = "";
        public string CameraMake { get; set; } = "";
        public string CameraModel { get; set; } = "";
        public string DateTaken { get; set; } = "";
        public string GPS_Latitude { get; set; } = "";
        public string GPS_Longitude { get; set; } = "";
        public string ISO { get; set; } = "";
        public string ExposureTime { get; set; } = "";
        public string FNumber { get; set; } = "";
        public string Flash { get; set; } = "";
        public string FocalLength { get; set; } = "";
        
        // Document Properties
        public string DocumentTitle { get; set; } = "";
        public string DocumentAuthor { get; set; } = "";
        public string DocumentSubject { get; set; } = "";
        public string DocumentKeywords { get; set; } = "";
        public string DocumentCategory { get; set; } = "";
        public string DocumentManager { get; set; } = "";
        public string DocumentCompany { get; set; } = "";
        public string LastSavedBy { get; set; } = "";
        public string ApplicationName { get; set; } = "";
        public int PageCount { get; set; }
        public int WordCount { get; set; }
        public int CharacterCount { get; set; }
        public string ContentType { get; set; } = "";
        
        // Audio/Video Metadata
        public string MediaDuration { get; set; } = "";
        public string AudioBitrate { get; set; } = "";
        public string VideoCodec { get; set; } = "";
        public string AudioCodec { get; set; } = "";
        public string FrameRate { get; set; } = "";
        public string AspectRatio { get; set; } = "";
        public string Album { get; set; } = "";
        public string Artist { get; set; } = "";
        public string Genre { get; set; } = "";
        public string Year { get; set; } = "";        public string Track { get; set; } = "";
          // Additional Image Properties
        public string BitDepth { get; set; } = "";
        public string HorizontalResolution { get; set; } = "";
        public string VerticalResolution { get; set; } = "";
        public string Orientation { get; set; } = "";
        public string FrameWidth { get; set; } = "";
        public string FrameHeight { get; set; } = "";
        
        // Language Property
        public string Language { get; set; } = "";
        
        // Archive Information
        public bool IsArchiveFile { get; set; }
        public int ArchiveFileCount { get; set; }
        public long ArchiveUncompressedSize { get; set; }
        
        // Additional Properties
        public string RegistryFileType { get; set; } = "";
        public string AlternateDataStreams { get; set; } = "";
        public string HardLinkCount { get; set; } = "";
        public string VolumeSerialNumber { get; set; } = "";
        public string DriveType { get; set; } = "";
        public string FileSystemType { get; set; } = "";
          // Comprehensive Properties Storage
        public string AllShellProperties { get; set; } = ""; // JSON string of all properties
        
        // Dynamic Properties Storage - This will store ALL discovered metadata properties
        public Dictionary<string, string> DynamicProperties { get; set; } = new Dictionary<string, string>();
    }
}