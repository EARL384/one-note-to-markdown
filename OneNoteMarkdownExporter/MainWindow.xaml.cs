using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using OneNoteMarkdownExporter.Models;
using OneNoteMarkdownExporter.Services;

namespace OneNoteMarkdownExporter
{
    public partial class MainWindow : Window
    {
        private readonly ExportService _exportService;
        private CancellationTokenSource? _cts;
        private const string NoFailuresMessage = "No failures.";
        private int _failureLogCount;

        public MainWindow()
        {
            InitializeComponent();
            _exportService = new ExportService();
            Loaded += MainWindow_Loaded;

            OneNoteItem.SelectionChanged += OnSelectionChanged;

            SetDefaultOutputPath();
            UpdateAssetOrganizationUi();
        }

        private void SetDefaultOutputPath()
        {
            try
            {
                var downloadsPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                downloadsPath = Path.Combine(downloadsPath, "Downloads", "OneNoteExport");

                if (!Directory.Exists(downloadsPath))
                {
                    Directory.CreateDirectory(downloadsPath);
                }

                OutputPathBox.Text = downloadsPath;
                AssetsPathBox.Text = AssetPathResolver.GetDefaultAssetsFolderPath(downloadsPath);
            }
            catch (Exception ex)
            {
                var fallbackPath = Path.Combine(Path.GetTempPath(), "OneNoteExport");
                if (!Directory.Exists(fallbackPath))
                {
                    Directory.CreateDirectory(fallbackPath);
                }

                OutputPathBox.Text = fallbackPath;
                AssetsPathBox.Text = AssetPathResolver.GetDefaultAssetsFolderPath(fallbackPath);
                Log($"Could not set Downloads folder, using temp folder: {ex.Message}");
            }
        }

        private void OnSelectionChanged(object? sender, EventArgs e)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(UpdateSelectionStatus);
            }
            else
            {
                UpdateSelectionStatus();
            }
        }

        private void UpdateSelectionStatus()
        {
            var items = NotebookTree.ItemsSource as List<OneNoteItem>;
            var selectedCount = items != null ? CountSelectedItems(items) : 0;

            if (selectedCount > 0)
            {
                SelectionStatusBorder.Background = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#D4EDDA"));
                SelectionStatusBorder.BorderBrush = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#C3E6CB"));
                SelectionStatusIcon.Text = "✓";
                SelectionStatusText.Text = $"{selectedCount} item{(selectedCount == 1 ? "" : "s")} selected for export";
                ExportButton.IsEnabled = true;
            }
            else
            {
                SelectionStatusBorder.Background = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#FFF3CD"));
                SelectionStatusBorder.BorderBrush = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#FFECB5"));
                SelectionStatusIcon.Text = "⚠";
                SelectionStatusText.Text = "Select notebooks, sections, or pages from the tree to export";
                ExportButton.IsEnabled = false;
            }
        }

        private int CountSelectedItems(List<OneNoteItem> items)
        {
            var count = 0;
            foreach (var item in items)
            {
                if (item.IsSelected) count++;
                count += CountSelectedItems(item.Children);
            }

            return count;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            LoadNotebooks();
        }

        private void LoadNotebooks()
        {
            try
            {
                var items = _exportService.GetNotebookHierarchy();
                NotebookTree.ItemsSource = items;
                Log("Notebooks loaded successfully.");
                UpdateSelectionStatus();
            }
            catch (Exception ex)
            {
                Log($"Error loading notebooks: {ex.Message}");
                System.Windows.MessageBox.Show("Error loading OneNote. Make sure OneNote Desktop is running.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            LoadNotebooks();
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var previousDefaultAssetsPath = AssetPathResolver.GetDefaultAssetsFolderPath(OutputPathBox.Text);
                    OutputPathBox.Text = dialog.SelectedPath;
                    if (GetSelectedAssetOrganizationMode() == AssetOrganizationMode.Centralized
                        && (string.IsNullOrWhiteSpace(AssetsPathBox.Text) || PathsEqual(AssetsPathBox.Text, previousDefaultAssetsPath)))
                    {
                        AssetsPathBox.Text = AssetPathResolver.GetDefaultAssetsFolderPath(dialog.SelectedPath);
                    }

                    UpdateAssetOrganizationUi();
                }
            }
        }

        private void AssetOrganizationBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            UpdateAssetOrganizationUi();
        }

        private AssetOrganizationMode GetSelectedAssetOrganizationMode()
        {
            return AssetOrganizationBox.SelectedIndex switch
            {
                1 => AssetOrganizationMode.Notebook,
                2 => AssetOrganizationMode.Section,
                3 => AssetOrganizationMode.Page,
                _ => AssetOrganizationMode.Centralized
            };
        }

        private void UpdateAssetOrganizationUi()
        {
            if (AssetsPathBox == null || BrowseAssetsButton == null || AssetOrganizationPreviewText == null)
            {
                return;
            }

            var mode = GetSelectedAssetOrganizationMode();
            var isCentralized = mode == AssetOrganizationMode.Centralized;

            AssetsPathBox.IsEnabled = isCentralized;
            BrowseAssetsButton.IsEnabled = isCentralized;
            AssetsPathBox.Opacity = isCentralized ? 1.0 : 0.55;
            BrowseAssetsButton.Opacity = isCentralized ? 1.0 : 0.55;

            AssetOrganizationPreviewText.Text = mode switch
            {
                AssetOrganizationMode.Centralized => "Assets are saved to one folder. You can choose the folder location.",
                AssetOrganizationMode.Notebook => "Each notebook folder gets a generated _assets_NotebookName folder.",
                AssetOrganizationMode.Section => "Each section folder gets a generated _assets_SectionName folder.",
                AssetOrganizationMode.Page => "Each page gets a generated _assets_PageName folder beside the Markdown file.",
                _ => string.Empty
            };
        }

        private void BrowseAssetsButton_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                if (!string.IsNullOrWhiteSpace(AssetsPathBox.Text) && Directory.Exists(AssetsPathBox.Text))
                {
                    dialog.SelectedPath = AssetsPathBox.Text;
                }

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    AssetsPathBox.Text = dialog.SelectedPath;
                }
            }
        }

        private async void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            var rootPath = OutputPathBox.Text;
            if (string.IsNullOrWhiteSpace(rootPath))
            {
                System.Windows.MessageBox.Show("Please select an output directory.");
                return;
            }

            var items = NotebookTree.ItemsSource as List<OneNoteItem>;
            if (items == null) return;

            var totalPages = ExportSelectionHelper.CountPagesToExport(items);
            if (totalPages == 0)
            {
                ExportProgressBar.IsIndeterminate = false;
                ExportProgressBar.Value = 0;
                SetExportStatus("No pages selected for export.", "#856404");
                Log("No pages selected for export.");
                return;
            }

            ExportButton.IsEnabled = false;
            StopButton.IsEnabled = true;
            NotebookTree.IsEnabled = false;
            RefreshButton.IsEnabled = false;
            ResetFailureLog();
            ExportProgressBar.Minimum = 0;
            ExportProgressBar.Maximum = totalPages;
            ExportProgressBar.Value = 0;
            ExportProgressBar.IsIndeterminate = false;
            SetExportStatus($"Exported 0 of {totalPages} pages...", "#0C5460");
            Log("Starting export...");

            var options = new ExportOptions
            {
                OutputPath = rootPath,
                AssetOrganizationMode = GetSelectedAssetOrganizationMode(),
                Overwrite = OverwriteExistingBox.IsChecked == true,
                ApplyLinting = LintMarkdownBox.IsChecked == true,
                PreserveDates = PreserveDatesBox.IsChecked == true,
                DateMetadataMode = YamlMetadataBox.IsChecked == true ? DateMetadataMode.Yaml : DateMetadataMode.None
            };

            if (options.AssetOrganizationMode == AssetOrganizationMode.Centralized)
            {
                options.AssetsFolderPath = AssetsPathBox.Text;
            }

            try
            {
                options.Validate();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "Assets Folder", MessageBoxButton.OK, MessageBoxImage.Error);
                ExportButton.IsEnabled = true;
                StopButton.IsEnabled = false;
                NotebookTree.IsEnabled = true;
                RefreshButton.IsEnabled = true;
                ExportProgressBar.IsIndeterminate = false;
                ExportProgressBar.Value = 0;
                LogGeneralFailure("Preparing assets folder", ex);
                ExportLogTabs.SelectedItem = FailureLogTab;
                SetExportStatus("Export failed.", "#721C24");
                return;
            }

            _cts = new CancellationTokenSource();
            var token = _cts.Token;
            ExportResult result;
            var generalFailureLogged = false;

            if (options.ApplyLinting && _exportService.IsMarkdownCliLinterAvailable)
            {
                Log("Markdown linting enabled (markdownlint-cli2)");
            }

            var progress = new Progress<ExportProgressUpdate>(HandleExportProgress);

            try
            {
                result = await _exportService.ExportSelectedAsync(items, options, progress, token);
            }
            catch (Exception ex)
            {
                result = new ExportResult { Error = ex.Message, TotalPages = totalPages };
                LogGeneralFailure("Export", ex);
                generalFailureLogged = true;
            }

            if (!generalFailureLogged && !string.IsNullOrEmpty(result.Error) && result.Failures.Count == 0)
            {
                LogGeneralFailure("Export", new Exception(result.Error));
            }

            ExportButton.IsEnabled = true;
            StopButton.IsEnabled = false;
            NotebookTree.IsEnabled = true;
            RefreshButton.IsEnabled = true;
            ExportProgressBar.IsIndeterminate = false;

            if (token.IsCancellationRequested)
            {
                ExportProgressBar.Value = result.ExportedPages;
                SetExportStatus($"Export stopped. {result.ExportedPages} of {totalPages} pages exported.", "#856404");
            }
            else if (!string.IsNullOrEmpty(result.Error) || result.FailedPages > 0)
            {
                ExportProgressBar.Value = result.ExportedPages;
                if (result.FailedPages > 0)
                {
                    SetExportStatus($"Export finished with errors. {result.ExportedPages} of {totalPages} pages exported, {result.FailedPages} failed.", "#721C24");
                }
                else
                {
                    SetExportStatus($"Export failed. {result.ExportedPages} of {totalPages} pages exported.", "#721C24");
                }
            }
            else
            {
                ExportProgressBar.Value = totalPages;
                SetExportStatus($"Export completed successfully. {result.ExportedPages} of {totalPages} pages exported.", "#155724");
            }

            if (_failureLogCount > 0)
            {
                ExportLogTabs.SelectedItem = FailureLogTab;
            }

            UpdateSelectionStatus();
            _cts = null;
        }

        private void StopButton_Click(object sender, RoutedEventArgs e)
        {
            if (_cts != null)
            {
                _cts.Cancel();
                Log("Stopping export... please wait for current operation to finish.");
                StopButton.IsEnabled = false;
            }
        }

        private void ConfigureLinting_Click(object sender, RoutedEventArgs e)
        {
            var configPath = Path.Combine(AppContext.BaseDirectory, "resources", ".markdownlint.json");
            if (File.Exists(configPath))
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = configPath,
                        UseShellExecute = true
                    });
                    Log($"Opened config file: {configPath}");
                }
                catch (Exception ex)
                {
                    Log($"Error opening config file: {ex.Message}");
                }
            }
            else
            {
                System.Windows.MessageBox.Show(
                    $"Config file not found at: {configPath}",
                    "Configuration",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        private void HandleExportProgress(ExportProgressUpdate update)
        {
            switch (update.Kind)
            {
                case ExportProgressKind.Warning:
                    Log(update.Message);
                    break;
                case ExportProgressKind.PageExported:
                    Log(update.Message);
                    UpdateExportProgress(update.ExportedPages, update.FailedPages, update.TotalPages);
                    break;
                case ExportProgressKind.PageFailed:
                    Log($"Export failure:\n{update.FailureDetails ?? update.Message}");
                    AppendFailureEntry(update.FailureDetails ?? update.Message);
                    UpdateExportProgress(update.ExportedPages, update.FailedPages, update.TotalPages);
                    break;
                default:
                    Log(update.Message);
                    break;
            }
        }

        private void UpdateExportProgress(int exportedPages, int failedPages, int totalPages)
        {
            ExportProgressBar.Value = exportedPages;
            var failedText = failedPages > 0 ? $", {failedPages} failed" : "";
            SetExportStatus($"Exported {exportedPages} of {totalPages} pages{failedText}...", "#0C5460");
        }

        private void SetExportStatus(string message, string colorHex)
        {
            ExportStatusText.Text = message;
            ExportStatusText.Foreground = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(colorHex));
        }

        private void ResetFailureLog()
        {
            _failureLogCount = 0;
            FailureLogTab.Header = "Failures (0)";
            FailureLogBox.Text = NoFailuresMessage;
            ExportLogTabs.SelectedItem = AllLogsTab;
        }

        private void LogGeneralFailure(string operation, Exception ex)
        {
            var failureDetails = ExportFailureFormatter.FormatGeneralFailure(operation, ex);
            Log($"Export failure:\n{failureDetails}");
            AppendFailureEntry(failureDetails);
        }

        private void AppendFailureEntry(string failureDetails)
        {
            if (!FailureLogBox.Dispatcher.CheckAccess())
            {
                FailureLogBox.Dispatcher.Invoke(() => AppendFailureEntry(failureDetails));
                return;
            }

            if (FailureLogBox.Text == NoFailuresMessage)
            {
                FailureLogBox.Clear();
            }

            _failureLogCount++;
            FailureLogTab.Header = $"Failures ({_failureLogCount})";
            FailureLogBox.AppendText($"{DateTime.Now:HH:mm:ss}: {failureDetails}{Environment.NewLine}{Environment.NewLine}");
            FailureLogBox.ScrollToEnd();
        }

        private static bool PathsEqual(string firstPath, string secondPath)
        {
            try
            {
                return string.Equals(
                    Path.GetFullPath(firstPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                    Path.GetFullPath(secondPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                    StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return string.Equals(firstPath, secondPath, StringComparison.OrdinalIgnoreCase);
            }
        }

        private void Log(string message)
        {
            if (LogBox.Dispatcher.CheckAccess())
            {
                LogBox.AppendText($"{DateTime.Now:HH:mm:ss}: {message}\n");
                LogBox.ScrollToEnd();
            }
            else
            {
                LogBox.Dispatcher.Invoke(() => Log(message));
            }
        }
    }
}
