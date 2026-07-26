using System;
using System.Collections.Generic;
using System.CommandLine;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OneNoteMarkdownExporter.Models;

namespace OneNoteMarkdownExporter.Services
{
    /// <summary>
    /// Handles command-line interface parsing and execution.
    /// </summary>
    public static class CliHandler
    {
        /// <summary>
        /// Parses command-line arguments and runs in CLI mode.
        /// </summary>
        /// <param name="args">Command-line arguments.</param>
        /// <returns>Exit code (0 for success, non-zero for failure).</returns>
        public static async Task<int> RunAsync(string[] args)
        {
            var rootCommand = BuildRootCommand();
            return await rootCommand.Parse(args).InvokeAsync();
        }

        /// <summary>
        /// Checks if CLI mode should be activated based on arguments.
        /// </summary>
        public static bool ShouldRunCli(string[] args)
        {
            // If there are any command-line arguments, run in CLI mode
            // Exceptions: arguments that VS/Windows might pass when launching GUI
            if (args.Length == 0) return false;

            // Check for known CLI flags
            var cliFlags = new[]
            {
                "--all", "--notebook", "--section", "--page", "--output", "-o",
                "--assets-folder", "--asset-organization", "--overwrite", "--no-lint", "--lint-config",
                "--no-preserve-dates", "--date-metadata", "--list", "--dry-run", "--verbose", "-v", "--quiet", "-q",
                "--help", "-h", "-?", "--version"
            };

            return args.Any(arg => cliFlags.Any(flag => 
                arg.StartsWith(flag, StringComparison.OrdinalIgnoreCase)));
        }

        private static RootCommand BuildRootCommand()
        {
            var rootCommand = new RootCommand("OneNote to Markdown Exporter - Export OneNote pages to Markdown files.")
            {
                TreatUnmatchedTokensAsErrors = true
            };

            // Options
            var allOption = new Option<bool>("--all")
            {
                Description = "Export all notebooks"
            };

            var notebookOption = new Option<string[]>("--notebook")
            {
                Description = "Export specific notebook(s) by name",
                AllowMultipleArgumentsPerToken = false
            };

            var sectionOption = new Option<string[]>("--section")
            {
                Description = "Export section(s) by path (e.g., 'Notebook/Section')",
                AllowMultipleArgumentsPerToken = false
            };

            var pageOption = new Option<string[]>("--page")
            {
                Description = "Export page(s) by ID",
                AllowMultipleArgumentsPerToken = false
            };

            var outputOption = new Option<string>("--output", "-o")
            {
                Description = "Output directory for exported files",
                DefaultValueFactory = _ => ExportOptions.GetDefaultOutputPath()
            };

            var assetsFolderOption = new Option<string?>("--assets-folder")
            {
                Description = "Path to folder for storing exported assets in centralized mode (default: <output>/_assets)"
            };

            var assetOrganizationOption = new Option<string>("--asset-organization")
            {
                Description = "Asset organization mode: centralized, notebook, section, or page",
                DefaultValueFactory = _ => "centralized"
            };

            var overwriteOption = new Option<bool>("--overwrite")
            {
                Description = "Overwrite existing files instead of creating numbered copies"
            };

            var noLintOption = new Option<bool>("--no-lint")
            {
                Description = "Disable Markdown linting (markdownlint-cli2)"
            };

            var lintConfigOption = new Option<string?>("--lint-config")
            {
                Description = "Path to custom .markdownlint.json configuration file"
            };

            var noPreserveDatesOption = new Option<bool>("--no-preserve-dates")
            {
                Description = "Do not preserve OneNote created/modified dates as file timestamps"
            };

            var dateMetadataOption = new Option<string>("--date-metadata")
            {
                Description = "Date metadata mode: none or yaml",
                DefaultValueFactory = _ => "none"
            };

            var listOption = new Option<bool>("--list")
            {
                Description = "List available notebooks, sections, and pages without exporting"
            };

            var dryRunOption = new Option<bool>("--dry-run")
            {
                Description = "Preview what would be exported without actually exporting"
            };

            var verboseOption = new Option<bool>("--verbose", "-v")
            {
                Description = "Show detailed output"
            };

            var quietOption = new Option<bool>("--quiet", "-q")
            {
                Description = "Show only errors"
            };

            // Add options to command
            rootCommand.Options.Add(allOption);
            rootCommand.Options.Add(notebookOption);
            rootCommand.Options.Add(sectionOption);
            rootCommand.Options.Add(pageOption);
            rootCommand.Options.Add(outputOption);
            rootCommand.Options.Add(assetsFolderOption);
            rootCommand.Options.Add(assetOrganizationOption);
            rootCommand.Options.Add(overwriteOption);
            rootCommand.Options.Add(noLintOption);
            rootCommand.Options.Add(lintConfigOption);
            rootCommand.Options.Add(noPreserveDatesOption);
            rootCommand.Options.Add(dateMetadataOption);
            rootCommand.Options.Add(listOption);
            rootCommand.Options.Add(dryRunOption);
            rootCommand.Options.Add(verboseOption);
            rootCommand.Options.Add(quietOption);

            rootCommand.SetAction(async (result, cancellationToken) =>
            {
                var assetOrganizationValue = result.GetValue(assetOrganizationOption);
                if (!TryParseAssetOrganizationMode(assetOrganizationValue, out var assetOrganizationMode))
                {
                    Console.Error.WriteLine($"Error: Unknown asset organization mode '{assetOrganizationValue}'. Use centralized, notebook, section, or page.");
                    return 1;
                }

                var dateMetadataValue = result.GetValue(dateMetadataOption);
                if (!TryParseDateMetadataMode(dateMetadataValue, out var dateMetadataMode))
                {
                    Console.Error.WriteLine($"Error: Unknown date metadata mode '{dateMetadataValue}'. Use none or yaml.");
                    return 1;
                }

                var options = new ExportOptions
                {
                    ExportAll = result.GetValue(allOption),
                    NotebookNames = result.GetValue(notebookOption)?.ToList(),
                    SectionPaths = result.GetValue(sectionOption)?.ToList(),
                    PageIds = result.GetValue(pageOption)?.ToList(),
                    OutputPath = result.GetValue(outputOption) ?? ExportOptions.GetDefaultOutputPath(),
                    AssetsFolderPath = result.GetValue(assetsFolderOption),
                    AssetOrganizationMode = assetOrganizationMode,
                    Overwrite = result.GetValue(overwriteOption),
                    ApplyLinting = !result.GetValue(noLintOption),
                    LintConfigPath = result.GetValue(lintConfigOption),
                    PreserveDates = !result.GetValue(noPreserveDatesOption),
                    DateMetadataMode = dateMetadataMode,
                    DryRun = result.GetValue(dryRunOption),
                    Verbose = result.GetValue(verboseOption),
                    Quiet = result.GetValue(quietOption)
                };

                var listMode = result.GetValue(listOption);

                return await ExecuteAsync(options, listMode, cancellationToken);
            });

            return rootCommand;
        }

        private static async Task<int> ExecuteAsync(ExportOptions options, bool listMode, CancellationToken cancellationToken)
        {
            var exportService = new ExportService();

            // Console progress reporter
            var progress = new Progress<ExportProgressUpdate>(update =>
            {
                if (!options.Quiet || update.Kind == ExportProgressKind.PageFailed || update.Message.Contains("Error") || update.Message.Contains("failed"))
                {
                    Console.WriteLine(update.Message);
                }
            });

            try
            {
                // List mode - just show hierarchy
                if (listMode)
                {
                    return ListHierarchy(exportService, options.Verbose);
                }

                try
                {
                    options.Validate();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Error: {ex.Message}");
                    return 1;
                }

                // Validate that we have selection criteria
                if (!options.HasSelectionCriteria())
                {
                    Console.Error.WriteLine("Error: No selection criteria specified.");
                    Console.Error.WriteLine("Use --all, --notebook, --section, or --page to specify what to export.");
                    Console.Error.WriteLine("Use --list to see available items.");
                    Console.Error.WriteLine("Use --help for more information.");
                    return 1;
                }

                // Report configuration
                if (!options.Quiet)
                {
                    Console.WriteLine("OneNote to Markdown Exporter");
                    Console.WriteLine("============================");
                    Console.WriteLine($"Output directory: {options.OutputPath}");
                    Console.WriteLine($"Asset organization: {FormatAssetOrganizationMode(options.AssetOrganizationMode)}");
                    if (options.AssetOrganizationMode == AssetOrganizationMode.Centralized)
                    {
                        Console.WriteLine($"Assets directory: {AssetPathResolver.ResolveAssetsFolderPath(options.OutputPath, options.AssetsFolderPath)}");
                    }
                    else
                    {
                        Console.WriteLine("Assets directory: generated per selected organization mode");
                    }
                    Console.WriteLine($"Overwrite: {(options.Overwrite ? "Yes" : "No")}");
                    var lintingStatus = !options.ApplyLinting
                        ? "Disabled"
                        : exportService.IsMarkdownCliLinterAvailable
                            ? "Enabled (markdownlint-cli2)"
                            : "Unavailable (export will continue without linting)";
                    Console.WriteLine($"Linting: {lintingStatus}");
                    Console.WriteLine($"Date preservation: {(options.PreserveDates ? "Enabled" : "Disabled")}");
                    Console.WriteLine($"Date metadata: {FormatDateMetadataMode(options.DateMetadataMode)}");
                    if (options.DryRun) Console.WriteLine("Mode: DRY RUN (no files will be created)");
                    Console.WriteLine();
                }

                // Run export
                var result = await exportService.ExportAsync(options, progress, cancellationToken);

                if (cancellationToken.IsCancellationRequested)
                {
                    return 130; // Standard exit code for Ctrl+C
                }

                if (!string.IsNullOrEmpty(result.Error))
                {
                    Console.Error.WriteLine($"Export error: {result.Error}");
                    return 1;
                }

                // Summary
                if (!options.Quiet && !options.DryRun)
                {
                    Console.WriteLine();
                    Console.WriteLine("Export Summary:");
                    Console.WriteLine($"  Pages exported: {result.ExportedPages}");
                    if (result.FailedPages > 0)
                    {
                        Console.WriteLine($"  Pages failed: {result.FailedPages}");
                    }
                    if (result.Warnings.Count > 0)
                    {
                        Console.WriteLine($"  Warnings: {result.Warnings.Count}");
                        foreach (var warning in result.Warnings)
                        {
                            Console.WriteLine($"  {warning}");
                        }
                    }
                }

                return result.FailedPages > 0 ? 1 : 0;
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Console.Error.WriteLine($"OneNote COM error: {ex.Message}");
                Console.Error.WriteLine("Make sure OneNote is installed and not running in a protected mode.");
                return 2;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
                if (options.Verbose)
                {
                    Console.Error.WriteLine(ex.StackTrace);
                }
                return 1;
            }
        }

        private static int ListHierarchy(ExportService exportService, bool verbose)
        {
            try
            {
                Console.WriteLine("OneNote Hierarchy");
                Console.WriteLine("=================");
                Console.WriteLine();

                var notebooks = exportService.GetNotebookHierarchy();

                if (notebooks.Count == 0)
                {
                    Console.WriteLine("No notebooks found.");
                    return 0;
                }

                foreach (var notebook in notebooks)
                {
                    PrintItem(notebook, "", verbose);
                }

                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error listing hierarchy: {ex.Message}");
                return 1;
            }
        }

        private static void PrintItem(OneNoteItem item, string indent, bool verbose)
        {
            var typeIcon = item.Type switch
            {
                OneNoteItemType.Notebook => "📓",
                OneNoteItemType.SectionGroup => "📁",
                OneNoteItemType.Section => "📄",
                OneNoteItemType.Page => "📝",
                _ => "❓"
            };

            var typeLabel = item.Type switch
            {
                OneNoteItemType.Notebook => "[Notebook]",
                OneNoteItemType.SectionGroup => "[SectionGroup]",
                OneNoteItemType.Section => "[Section]",
                OneNoteItemType.Page => "[Page]",
                _ => "[Unknown]"
            };

            if (verbose)
            {
                Console.WriteLine($"{indent}{typeIcon} {item.Name} {typeLabel}");
                Console.WriteLine($"{indent}   ID: {item.Id}");
            }
            else
            {
                Console.WriteLine($"{indent}{typeIcon} {item.Name}");
            }

            foreach (var child in item.Children)
            {
                PrintItem(child, indent + "  ", verbose);
            }
        }

        public static bool TryParseAssetOrganizationMode(string? value, out AssetOrganizationMode mode)
        {
            switch (value?.Trim().ToLowerInvariant())
            {
                case null:
                case "":
                case "centralized":
                    mode = AssetOrganizationMode.Centralized;
                    return true;
                case "notebook":
                    mode = AssetOrganizationMode.Notebook;
                    return true;
                case "section":
                    mode = AssetOrganizationMode.Section;
                    return true;
                case "page":
                    mode = AssetOrganizationMode.Page;
                    return true;
                default:
                    mode = AssetOrganizationMode.Centralized;
                    return false;
            }
        }

        public static bool TryParseDateMetadataMode(string? value, out DateMetadataMode mode)
        {
            switch (value?.Trim().ToLowerInvariant())
            {
                case null:
                case "":
                case "none":
                    mode = DateMetadataMode.None;
                    return true;
                case "yaml":
                    mode = DateMetadataMode.Yaml;
                    return true;
                default:
                    mode = DateMetadataMode.None;
                    return false;
            }
        }

        private static string FormatAssetOrganizationMode(AssetOrganizationMode mode)
        {
            return mode switch
            {
                AssetOrganizationMode.Centralized => "centralized",
                AssetOrganizationMode.Notebook => "notebook",
                AssetOrganizationMode.Section => "section",
                AssetOrganizationMode.Page => "page",
                _ => mode.ToString()
            };
        }

        private static string FormatDateMetadataMode(DateMetadataMode mode)
        {
            return mode switch
            {
                DateMetadataMode.None => "none",
                DateMetadataMode.Yaml => "yaml",
                _ => mode.ToString()
            };
        }
    }
}
