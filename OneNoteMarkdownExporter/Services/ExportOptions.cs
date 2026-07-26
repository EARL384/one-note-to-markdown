using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace OneNoteMarkdownExporter.Services
{
    public enum AssetOrganizationMode
    {
        Centralized,
        Notebook,
        Section,
        Page
    }

    public enum DateMetadataMode
    {
        None,
        Yaml
    }

    /// <summary>
    /// Options for controlling the export process.
    /// Used by both GUI and CLI modes.
    /// </summary>
    public class ExportOptions
    {
        /// <summary>
        /// Root directory where exported files will be saved.
        /// </summary>
        public string OutputPath { get; set; } = string.Empty;

        /// <summary>
        /// Folder where exported assets will be saved. If null/empty, uses OutputPath\_assets.
        /// Relative paths are resolved from the output directory.
        /// </summary>
        public string? AssetsFolderPath { get; set; }

        /// <summary>
        /// Controls how exported assets are grouped on disk.
        /// </summary>
        public AssetOrganizationMode AssetOrganizationMode { get; set; } = AssetOrganizationMode.Centralized;

        /// <summary>
        /// If true, overwrite existing files. If false, create numbered copies.
        /// </summary>
        public bool Overwrite { get; set; } = false;

        /// <summary>
        /// If true, apply Markdown linting/formatting to exported content using markdownlint-cli2.
        /// Linting errors are logged but do not prevent export.
        /// </summary>
        public bool ApplyLinting { get; set; } = true;

        /// <summary>
        /// Path to custom .markdownlint.json config file.
        /// If null/empty, uses the default bundled config.
        /// </summary>
        public string? LintConfigPath { get; set; }

        // Selection options - at least one should be set for CLI mode

        /// <summary>
        /// If true, export all notebooks.
        /// </summary>
        public bool ExportAll { get; set; } = false;

        /// <summary>
        /// List of notebook names to export. Case-insensitive matching.
        /// </summary>
        public List<string>? NotebookNames { get; set; }

        /// <summary>
        /// List of section paths to export (format: "NotebookName/SectionName" or "NotebookName/SectionGroupName/SectionName").
        /// </summary>
        public List<string>? SectionPaths { get; set; }

        /// <summary>
        /// List of specific page IDs to export.
        /// </summary>
        public List<string>? PageIds { get; set; }

        /// <summary>
        /// If true, show what would be exported without actually exporting.
        /// </summary>
        public bool DryRun { get; set; } = false;

        /// <summary>
        /// If true, output verbose logging.
        /// </summary>
        public bool Verbose { get; set; } = false;

        /// <summary>
        /// If true, suppress all output except errors.
        /// </summary>
        public bool Quiet { get; set; } = false;

        public bool PreserveDates { get; set; } = true;
        public DateMetadataMode DateMetadataMode { get; set; } = DateMetadataMode.None;

        /// <summary>
        /// Creates default export options.
        /// </summary>
        public static ExportOptions CreateDefault()
        {
            return new ExportOptions
            {
                ApplyLinting = true,
                Overwrite = false
            };
        }

        /// <summary>
        /// Gets the default output path (Downloads\OneNoteExport).
        /// </summary>
        public static string GetDefaultOutputPath()
        {
            string downloadsPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.UserProfile);
            return System.IO.Path.Combine(downloadsPath, "Downloads", "OneNoteExport");
        }

        /// <summary>
        /// Checks if any selection criteria is specified.
        /// </summary>
        public bool HasSelectionCriteria()
        {
            return ExportAll ||
                   (NotebookNames != null && NotebookNames.Count > 0) ||
                   (SectionPaths != null && SectionPaths.Count > 0) ||
                   (PageIds != null && PageIds.Count > 0);
        }

        public void Validate()
        {
            if (AssetOrganizationMode != AssetOrganizationMode.Centralized && !string.IsNullOrWhiteSpace(AssetsFolderPath))
            {
                throw new InvalidOperationException("Custom assets folder paths are only supported when asset organization is centralized.");
            }

            if (!ApplyLinting || string.IsNullOrWhiteSpace(LintConfigPath))
            {
                return;
            }

            LintConfigPath = Path.GetFullPath(LintConfigPath);
            if (!Path.GetExtension(LintConfigPath).Equals(".json", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Markdown lint configuration must be a JSON file.");
            }

            if (!File.Exists(LintConfigPath))
            {
                throw new InvalidOperationException($"Markdown lint configuration not found at: {LintConfigPath}");
            }

            try
            {
                using var document = JsonDocument.Parse(File.ReadAllText(LintConfigPath));
            }
            catch (JsonException ex)
            {
                throw new InvalidOperationException($"Invalid Markdown lint configuration '{LintConfigPath}': {ex.Message}", ex);
            }
        }
    }
}
