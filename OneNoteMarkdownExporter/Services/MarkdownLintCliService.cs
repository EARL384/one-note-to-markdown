using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace OneNoteMarkdownExporter.Services
{
    /// <summary>
    /// Service for running markdownlint-cli2 using the bundled Node.js runtime.
    /// This provides full markdownlint compatibility without requiring users to install Node.js.
    /// </summary>
    public class MarkdownLintCliService : IMarkdownLintService
    {
        private readonly string _nodeExePath;
        private readonly string _markdownLintPath;
        private readonly string _configPath;
        private bool _isAvailable;

        public bool IsAvailable => _isAvailable;
        public string UnavailableReason { get; private set; } = "";

        public MarkdownLintCliService(string? resourcesDirectory = null)
        {
            var resourcesDir = resourcesDirectory ?? Path.Combine(AppContext.BaseDirectory, "resources");

            _nodeExePath = Path.Combine(resourcesDir, "node.exe");
            _markdownLintPath = Path.Combine(resourcesDir, "markdownlint-cli2.mjs");
            _configPath = Path.Combine(resourcesDir, ".markdownlint.json");

            CheckAvailability();
        }

        private void CheckAvailability()
        {
            if (!File.Exists(_nodeExePath))
            {
                _isAvailable = false;
                UnavailableReason = $"node.exe not found at: {_nodeExePath}";
                return;
            }

            if (!File.Exists(_markdownLintPath))
            {
                _isAvailable = false;
                UnavailableReason = $"markdownlint-cli2 bundle not found at: {_markdownLintPath}";
                return;
            }

            if (!File.Exists(_configPath))
            {
                _isAvailable = false;
                UnavailableReason = $"Markdown lint configuration not found at: {_configPath}";
                return;
            }

            try
            {
                ValidateConfigFile(_configPath);
            }
            catch (Exception ex)
            {
                _isAvailable = false;
                UnavailableReason = ex.Message;
                return;
            }

            _isAvailable = true;
            UnavailableReason = "";
        }

        /// <summary>
        /// Lints and fixes a markdown file in place using markdownlint-cli2.
        /// </summary>
        /// <param name="filePath">Path to the markdown file to lint.</param>
        /// <param name="configPath">Optional path to a custom JSON configuration file.</param>
        /// <returns>Result containing success status and any output messages.</returns>
        public async Task<LintResult> LintFileAsync(string filePath, string? configPath = null)
        {
            if (!_isAvailable)
            {
                return new LintResult
                {
                    Success = false,
                    ErrorMessage = UnavailableReason
                };
            }

            if (!File.Exists(filePath))
            {
                return new LintResult
                {
                    Success = false,
                    ErrorMessage = $"File not found: {filePath}"
                };
            }

            try
            {
                var selectedConfigPath = string.IsNullOrWhiteSpace(configPath)
                    ? _configPath
                    : Path.GetFullPath(configPath);
                ValidateConfigFile(selectedConfigPath);

                var workingDirectory = Path.GetDirectoryName(Path.GetFullPath(filePath)) ?? Path.GetTempPath();
                var literalFileArgument = $":./{Path.GetFileName(filePath)}";
                var startInfo = new ProcessStartInfo
                {
                    FileName = _nodeExePath,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WorkingDirectory = workingDirectory
                };
                startInfo.ArgumentList.Add(_markdownLintPath);
                startInfo.ArgumentList.Add(literalFileArgument);
                startInfo.ArgumentList.Add("--config");
                startInfo.ArgumentList.Add(selectedConfigPath);
                startInfo.ArgumentList.Add("--fix");

                using var process = new Process
                {
                    StartInfo = startInfo
                };

                process.Start();
                var outputTask = process.StandardOutput.ReadToEndAsync();
                var errorTask = process.StandardError.ReadToEndAsync();
                await process.WaitForExitAsync();

                var output = await outputTask;
                var error = await errorTask;

                // Exit code 0 = success, 1 = lint errors found (but --fix applied what it could)
                // Exit code 2 = execution or configuration error
                if (process.ExitCode == 2)
                {
                    return new LintResult
                    {
                        Success = false,
                        ErrorMessage = GetFirstErrorLine(error, output),
                        Output = output
                    };
                }

                if (process.ExitCode != 0 && process.ExitCode != 1)
                {
                    return new LintResult
                    {
                        Success = false,
                        ErrorMessage = $"markdownlint-cli2 exited with unexpected code {process.ExitCode}.",
                        Output = output
                    };
                }

                return new LintResult
                {
                    Success = true,
                    Output = output,
                    WarningMessage = process.ExitCode == 1
                        ? $"{output.Trim()}{Environment.NewLine}{error.Trim()}".Trim()
                        : error.Trim()
                };
            }
            catch (Exception ex)
            {
                return new LintResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to run markdownlint-cli2: {ex.Message}"
                };
            }
        }

        /// <summary>
        /// Lints markdown content (not a file) by writing to a temp file, linting, and reading back.
        /// </summary>
        /// <param name="markdown">The markdown content to lint.</param>
        /// <returns>The linted markdown content.</returns>
        public async Task<LintResult> LintContentAsync(string markdown, string? configPath = null)
        {
            if (!_isAvailable)
            {
                return new LintResult
                {
                    Success = false,
                    Content = markdown,
                    ErrorMessage = UnavailableReason
                };
            }

            var tempFile = Path.Combine(Path.GetTempPath(), $"mdlint_{Guid.NewGuid():N}.md");

            try
            {
                await File.WriteAllTextAsync(tempFile, markdown, new UTF8Encoding(false));

                var result = await LintFileAsync(tempFile, configPath);

                if (result.Success)
                {
                    result.Content = await File.ReadAllTextAsync(tempFile);
                }
                else
                {
                    result.Content = markdown;
                }

                return result;
            }
            catch (Exception ex)
            {
                return new LintResult
                {
                    Success = false,
                    Content = markdown,
                    ErrorMessage = $"Failed to run markdownlint-cli2: {ex.Message}"
                };
            }
            finally
            {
                try
                {
                    if (File.Exists(tempFile))
                        File.Delete(tempFile);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }

        private static void ValidateConfigFile(string configPath)
        {
            if (!File.Exists(configPath))
            {
                throw new FileNotFoundException($"Markdown lint configuration not found at: {configPath}", configPath);
            }

            if (!Path.GetExtension(configPath).Equals(".json", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("Markdown lint configuration must be a JSON file.");
            }

            try
            {
                using var document = JsonDocument.Parse(File.ReadAllText(configPath));
            }
            catch (JsonException ex)
            {
                throw new InvalidDataException($"Invalid Markdown lint configuration '{configPath}': {ex.Message}", ex);
            }
        }

        private static string GetFirstErrorLine(string error, string output)
        {
            var message = string.IsNullOrWhiteSpace(error) ? output : error;
            using var reader = new StringReader(message);
            return reader.ReadLine()?.Trim() ?? "markdownlint-cli2 failed without an error message.";
        }

        /// <summary>
        /// Synchronous version of LintContentAsync for compatibility.
        /// </summary>
        public LintResult LintContent(string markdown, string? configPath = null)
        {
            return LintContentAsync(markdown, configPath).GetAwaiter().GetResult();
        }
    }

    public class LintResult
    {
        public bool Success { get; set; }
        public string Content { get; set; } = "";
        public string Output { get; set; } = "";
        public string ErrorMessage { get; set; } = "";
        public string WarningMessage { get; set; } = "";
    }
}
