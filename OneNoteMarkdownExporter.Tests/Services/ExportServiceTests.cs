using System.IO;
using FluentAssertions;
using OneNoteMarkdownExporter.Models;
using OneNoteMarkdownExporter.Services;
using Xunit;

namespace OneNoteMarkdownExporter.Tests.Services;

public class ExportServiceTests
{
    [Theory]
    [InlineData(AssetOrganizationMode.Centralized, @"..\..\_assets", @"_assets")]
    [InlineData(AssetOrganizationMode.Notebook, @"..\_assets_ProjectNotebook", @"Project Notebook\_assets_ProjectNotebook")]
    [InlineData(AssetOrganizationMode.Section, @"_assets_MeetingNotes", @"Project Notebook\Meeting Notes\_assets_MeetingNotes")]
    [InlineData(AssetOrganizationMode.Page, @"_assets_PlanningPage", @"Project Notebook\Meeting Notes\_assets_PlanningPage")]
    public async Task ExportSelectedAsync_WithAssetOrganizationMode_UsesExpectedAssetsFolderAndRelativePath(
        AssetOrganizationMode mode,
        string expectedRelativeAssetsPath,
        string expectedAssetsRelativeFolder)
    {
        var outputPath = CreateTempExportPath();
        var source = new FakeOneNoteExportSource();
        var converter = new RecordingMarkdownConverter();
        var linter = new PassthroughMarkdownLintService();
        var service = new ExportService(source, converter, linter);
        var notebooks = CreateSelectedHierarchy();
        var options = new ExportOptions
        {
            OutputPath = outputPath,
            AssetOrganizationMode = mode,
            ApplyLinting = false,
            Overwrite = true
        };

        try
        {
            var result = await service.ExportSelectedAsync(notebooks, options);

            result.Success.Should().BeTrue();
            result.ExportedPages.Should().Be(1);
            converter.Calls.Should().ContainSingle();
            converter.Calls[0].AssetsFolder.Should().Be(Path.Combine(outputPath, expectedAssetsRelativeFolder));
            converter.Calls[0].RelativeAssetsPath.Should().Be(expectedRelativeAssetsPath.Replace("\\", "/"));
            File.ReadAllText(Path.Combine(outputPath, "Project Notebook", "Meeting Notes", "Planning Page.md"))
                .Should().Be($"asset:{expectedRelativeAssetsPath.Replace("\\", "/")}");
            Directory.Exists(Path.Combine(outputPath, expectedAssetsRelativeFolder)).Should().BeFalse();
        }
        finally
        {
            if (Directory.Exists(outputPath)) Directory.Delete(outputPath, true);
        }
    }

    [Fact]
    public async Task ExportSelectedAsync_WithCentralizedCustomAssetsFolder_UsesCustomFolder()
    {
        var tempRoot = CreateTempExportPath();
        var outputPath = Path.Combine(tempRoot, "export");
        var assetsPath = Path.Combine(tempRoot, "shared-assets");
        var source = new FakeOneNoteExportSource();
        var converter = new RecordingMarkdownConverter();
        var linter = new PassthroughMarkdownLintService();
        var service = new ExportService(source, converter, linter);
        var notebooks = CreateSelectedHierarchy();
        var options = new ExportOptions
        {
            OutputPath = outputPath,
            AssetsFolderPath = assetsPath,
            AssetOrganizationMode = AssetOrganizationMode.Centralized,
            ApplyLinting = false,
            Overwrite = true
        };

        try
        {
            var result = await service.ExportSelectedAsync(notebooks, options);

            result.Success.Should().BeTrue();
            converter.Calls.Should().ContainSingle();
            converter.Calls[0].AssetsFolder.Should().Be(assetsPath);
            converter.Calls[0].RelativeAssetsPath.Should().Be("../../../shared-assets");
            Directory.Exists(assetsPath).Should().BeFalse();
        }
        finally
        {
            if (Directory.Exists(tempRoot)) Directory.Delete(tempRoot, true);
        }
    }

    [Theory]
    [InlineData(AssetOrganizationMode.Centralized, @"_assets")]
    [InlineData(AssetOrganizationMode.Notebook, @"Project Notebook\_assets_ProjectNotebook")]
    [InlineData(AssetOrganizationMode.Section, @"Project Notebook\Meeting Notes\_assets_MeetingNotes")]
    [InlineData(AssetOrganizationMode.Page, @"Project Notebook\Meeting Notes\_assets_PlanningPage")]
    public async Task ExportSelectedAsync_WithTextOnlyPage_DoesNotCreateAssetsFolder(AssetOrganizationMode mode, string expectedAssetsRelativeFolder)
    {
        var outputPath = CreateTempExportPath();
        var source = new FakeOneNoteExportSource();
        var converter = new RecordingMarkdownConverter();
        var linter = new PassthroughMarkdownLintService();
        var service = new ExportService(source, converter, linter);
        var notebooks = CreateSelectedHierarchy();
        var options = new ExportOptions
        {
            OutputPath = outputPath,
            AssetOrganizationMode = mode,
            ApplyLinting = false,
            Overwrite = true
        };

        try
        {
            var result = await service.ExportSelectedAsync(notebooks, options);

            result.Success.Should().BeTrue();
            Directory.Exists(Path.Combine(outputPath, expectedAssetsRelativeFolder)).Should().BeFalse();
        }
        finally
        {
            if (Directory.Exists(outputPath)) Directory.Delete(outputPath, true);
        }
    }

    [Fact]
    public async Task ExportSelectedAsync_WhenConverterWritesAsset_CreatesAssetsFolder()
    {
        var outputPath = CreateTempExportPath();
        var source = new FakeOneNoteExportSource();
        var converter = new AssetWritingMarkdownConverter();
        var linter = new PassthroughMarkdownLintService();
        var service = new ExportService(source, converter, linter);
        var notebooks = CreateSelectedHierarchy();
        var options = new ExportOptions
        {
            OutputPath = outputPath,
            AssetOrganizationMode = AssetOrganizationMode.Page,
            ApplyLinting = false,
            Overwrite = true
        };
        var expectedAssetsFolder = Path.Combine(outputPath, "Project Notebook", "Meeting Notes", "_assets_PlanningPage");

        try
        {
            var result = await service.ExportSelectedAsync(notebooks, options);

            result.Success.Should().BeTrue();
            Directory.Exists(expectedAssetsFolder).Should().BeTrue();
            File.Exists(Path.Combine(expectedAssetsFolder, "image_0001.png")).Should().BeTrue();
        }
        finally
        {
            if (Directory.Exists(outputPath)) Directory.Delete(outputPath, true);
        }
    }

    [Fact]
    public async Task ExportSelectedAsync_WithExistingFileAtGeneratedAssetsPath_FailsPageClearly()
    {
        var outputPath = CreateTempExportPath();
        var assetsFilePath = Path.Combine(outputPath, "Project Notebook", "Meeting Notes", "_assets_PlanningPage");
        var source = new FakeOneNoteExportSource();
        var converter = new RecordingMarkdownConverter();
        var linter = new PassthroughMarkdownLintService();
        var service = new ExportService(source, converter, linter);
        var notebooks = CreateSelectedHierarchy();
        var options = new ExportOptions
        {
            OutputPath = outputPath,
            AssetOrganizationMode = AssetOrganizationMode.Page,
            ApplyLinting = false,
            Overwrite = true
        };

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(assetsFilePath)!);
            File.WriteAllText(assetsFilePath, "not a folder");

            var result = await service.ExportSelectedAsync(notebooks, options);

            result.Success.Should().BeFalse();
            result.FailedPages.Should().Be(1);
            result.Failures.Should().ContainSingle()
                .Which.Should().Contain("Assets folder path points to an existing file:");
        }
        finally
        {
            if (Directory.Exists(outputPath)) Directory.Delete(outputPath, true);
        }
    }

    [Fact]
    public async Task ExportSelectedAsync_WithExistingFileAtCustomCentralizedAssetsPath_FailsClearly()
    {
        var tempRoot = CreateTempExportPath();
        var outputPath = Path.Combine(tempRoot, "export");
        var assetsFilePath = Path.Combine(tempRoot, "custom-assets");
        var source = new FakeOneNoteExportSource();
        var converter = new RecordingMarkdownConverter();
        var linter = new PassthroughMarkdownLintService();
        var service = new ExportService(source, converter, linter);
        var notebooks = CreateSelectedHierarchy();
        var options = new ExportOptions
        {
            OutputPath = outputPath,
            AssetsFolderPath = assetsFilePath,
            AssetOrganizationMode = AssetOrganizationMode.Centralized,
            ApplyLinting = false,
            Overwrite = true
        };

        try
        {
            Directory.CreateDirectory(tempRoot);
            File.WriteAllText(assetsFilePath, "not a folder");

            var result = await service.ExportSelectedAsync(notebooks, options);

            result.Success.Should().BeFalse();
            result.Error.Should().Contain("Assets folder path points to an existing file:");
        }
        finally
        {
            if (Directory.Exists(tempRoot)) Directory.Delete(tempRoot, true);
        }
    }

    [Fact]
    public async Task ExportSelectedAsync_WithPreserveDatesEnabled_AppliesTimestampsAfterWrite()
    {
        var outputPath = CreateTempExportPath();
        var timestampService = new RecordingFileTimestampService();
        var service = new ExportService(new FakeOneNoteExportSource(), new RecordingMarkdownConverter(), new PassthroughMarkdownLintService(), timestampService, new PassthroughYamlFrontMatterService());
        var notebooks = CreateSelectedHierarchy(created: new DateTimeOffset(2024, 1, 15, 10, 30, 0, TimeSpan.Zero), modified: new DateTimeOffset(2024, 2, 20, 14, 45, 0, TimeSpan.Zero));
        var options = new ExportOptions { OutputPath = outputPath, ApplyLinting = false, Overwrite = true };
        var expectedFilePath = Path.Combine(outputPath, "Project Notebook", "Meeting Notes", "Planning Page.md");

        try
        {
            var result = await service.ExportSelectedAsync(notebooks, options);

            result.Success.Should().BeTrue();
            timestampService.Calls.Should().ContainSingle();
            timestampService.Calls[0].FilePath.Should().Be(expectedFilePath);
            timestampService.Calls[0].Created.Should().Be(notebooks[0].Children[0].Children[0].CreatedTime);
            timestampService.Calls[0].Modified.Should().Be(notebooks[0].Children[0].Children[0].LastModifiedTime);
            File.Exists(expectedFilePath).Should().BeTrue();
        }
        finally
        {
            if (Directory.Exists(outputPath)) Directory.Delete(outputPath, true);
        }
    }

    [Fact]
    public async Task ExportSelectedAsync_WithPreserveDatesDisabled_DoesNotApplyTimestamps()
    {
        var outputPath = CreateTempExportPath();
        var timestampService = new RecordingFileTimestampService();
        var service = new ExportService(new FakeOneNoteExportSource(), new RecordingMarkdownConverter(), new PassthroughMarkdownLintService(), timestampService, new PassthroughYamlFrontMatterService());
        var notebooks = CreateSelectedHierarchy(created: new DateTimeOffset(2024, 1, 15, 10, 30, 0, TimeSpan.Zero), modified: new DateTimeOffset(2024, 2, 20, 14, 45, 0, TimeSpan.Zero));
        var options = new ExportOptions { OutputPath = outputPath, PreserveDates = false, ApplyLinting = false, Overwrite = true };

        try
        {
            var result = await service.ExportSelectedAsync(notebooks, options);

            result.Success.Should().BeTrue();
            timestampService.Calls.Should().BeEmpty();
        }
        finally
        {
            if (Directory.Exists(outputPath)) Directory.Delete(outputPath, true);
        }
    }

    [Fact]
    public async Task ExportSelectedAsync_WithNoPageDates_DoesNotApplyTimestamps()
    {
        var outputPath = CreateTempExportPath();
        var timestampService = new RecordingFileTimestampService();
        var service = new ExportService(new FakeOneNoteExportSource(), new RecordingMarkdownConverter(), new PassthroughMarkdownLintService(), timestampService, new PassthroughYamlFrontMatterService());
        var notebooks = CreateSelectedHierarchy();
        var options = new ExportOptions { OutputPath = outputPath, ApplyLinting = false, Overwrite = true };

        try
        {
            var result = await service.ExportSelectedAsync(notebooks, options);

            result.Success.Should().BeTrue();
            timestampService.Calls.Should().BeEmpty();
        }
        finally
        {
            if (Directory.Exists(outputPath)) Directory.Delete(outputPath, true);
        }
    }

    [Fact]
    public async Task ExportSelectedAsync_WhenTimestampServiceThrows_AddsWarningAndExportsPage()
    {
        var outputPath = CreateTempExportPath();
        var timestampService = new ThrowingFileTimestampService();
        var service = new ExportService(new FakeOneNoteExportSource(), new RecordingMarkdownConverter(), new PassthroughMarkdownLintService(), timestampService, new PassthroughYamlFrontMatterService());
        var notebooks = CreateSelectedHierarchy(modified: new DateTimeOffset(2024, 2, 20, 14, 45, 0, TimeSpan.Zero));
        var options = new ExportOptions { OutputPath = outputPath, ApplyLinting = false, Overwrite = true };

        try
        {
            var result = await service.ExportSelectedAsync(notebooks, options);

            result.Success.Should().BeTrue();
            result.ExportedPages.Should().Be(1);
            result.FailedPages.Should().Be(0);
            result.Warnings.Should().ContainSingle().Which.Should().Contain("Warning: Could not preserve dates for 'Planning Page': timestamp failed");
        }
        finally
        {
            if (Directory.Exists(outputPath)) Directory.Delete(outputPath, true);
        }
    }

    [Fact]
    public async Task ExportSelectedAsync_WithDateMetadataNone_DoesNotAddYaml()
    {
        var outputPath = CreateTempExportPath();
        var yamlService = new RecordingYamlFrontMatterService();
        var service = new ExportService(new FakeOneNoteExportSource(), new RecordingMarkdownConverter(), new PassthroughMarkdownLintService(), new RecordingFileTimestampService(), yamlService);
        var notebooks = CreateSelectedHierarchy(modified: new DateTimeOffset(2024, 2, 20, 14, 45, 0, TimeSpan.Zero));
        var options = new ExportOptions { OutputPath = outputPath, DateMetadataMode = DateMetadataMode.None, ApplyLinting = false, Overwrite = true };

        try
        {
            var result = await service.ExportSelectedAsync(notebooks, options);

            result.Success.Should().BeTrue();
            yamlService.Calls.Should().Be(0);
            File.ReadAllText(Path.Combine(outputPath, "Project Notebook", "Meeting Notes", "Planning Page.md")).Should().NotStartWith("---");
        }
        finally
        {
            if (Directory.Exists(outputPath)) Directory.Delete(outputPath, true);
        }
    }

    [Fact]
    public async Task ExportSelectedAsync_WithDateMetadataYaml_AddsYamlBeforeLinting()
    {
        var outputPath = CreateTempExportPath();
        var yamlService = new RecordingYamlFrontMatterService();
        var linter = new RecordingMarkdownLintService();
        var service = new ExportService(new FakeOneNoteExportSource(), new RecordingMarkdownConverter(), linter, new RecordingFileTimestampService(), yamlService);
        var notebooks = CreateSelectedHierarchy(modified: new DateTimeOffset(2024, 2, 20, 14, 45, 0, TimeSpan.Zero));
        var options = new ExportOptions { OutputPath = outputPath, DateMetadataMode = DateMetadataMode.Yaml, ApplyLinting = true, Overwrite = true };

        try
        {
            var result = await service.ExportSelectedAsync(notebooks, options);

            result.Success.Should().BeTrue();
            yamlService.Calls.Should().Be(1);
            linter.Inputs.Should().ContainSingle().Which.Should().StartWith("---\nmetadata: true\n---\n\n");
        }
        finally
        {
            if (Directory.Exists(outputPath)) Directory.Delete(outputPath, true);
        }
    }

    [Fact]
    public async Task ExportSelectedAsync_WithCustomLintConfig_ForwardsNormalizedPath()
    {
        var outputPath = CreateTempExportPath();
        var configPath = Path.Combine(outputPath, "custom.markdownlint.json");
        Directory.CreateDirectory(outputPath);
        File.WriteAllText(configPath, "{\"default\":true}");
        var linter = new RecordingMarkdownLintService();
        var service = new ExportService(new FakeOneNoteExportSource(), new RecordingMarkdownConverter(), linter);
        var options = new ExportOptions
        {
            OutputPath = outputPath,
            LintConfigPath = configPath,
            ApplyLinting = true,
            Overwrite = true
        };

        try
        {
            var result = await service.ExportSelectedAsync(CreateSelectedHierarchy(), options);

            result.Success.Should().BeTrue();
            linter.ConfigPaths.Should().ContainSingle().Which.Should().Be(Path.GetFullPath(configPath));
        }
        finally
        {
            if (Directory.Exists(outputPath)) Directory.Delete(outputPath, true);
        }
    }

    [Fact]
    public async Task ExportSelectedAsync_WhenLintRuntimeUnavailable_WarnsOnceAndExportsOriginalContent()
    {
        var outputPath = CreateTempExportPath();
        var linter = new StubMarkdownLintService
        {
            IsAvailable = false,
            UnavailableReason = "runtime missing"
        };
        var service = new ExportService(new FakeOneNoteExportSource(), new RecordingMarkdownConverter(), linter);
        var options = new ExportOptions { OutputPath = outputPath, ApplyLinting = true, Overwrite = true };

        try
        {
            var result = await service.ExportSelectedAsync(CreateSelectedHierarchy(), options);
            var markdownPath = Path.Combine(outputPath, "Project Notebook", "Meeting Notes", "Planning Page.md");

            result.Success.Should().BeTrue();
            result.Warnings.Should().ContainSingle().Which.Should().Contain("Markdown linting is unavailable: runtime missing");
            linter.Calls.Should().Be(0);
            File.ReadAllText(markdownPath).Should().StartWith("asset:");
        }
        finally
        {
            if (Directory.Exists(outputPath)) Directory.Delete(outputPath, true);
        }
    }

    [Fact]
    public async Task ExportSelectedAsync_WhenLintingFails_WarnsAndExportsOriginalContent()
    {
        var outputPath = CreateTempExportPath();
        var linter = new StubMarkdownLintService
        {
            IsAvailable = true,
            Result = new LintResult { Success = false, Content = "discarded", ErrorMessage = "lint failed" }
        };
        var service = new ExportService(new FakeOneNoteExportSource(), new RecordingMarkdownConverter(), linter);
        var options = new ExportOptions { OutputPath = outputPath, ApplyLinting = true, Overwrite = true };

        try
        {
            var result = await service.ExportSelectedAsync(CreateSelectedHierarchy(), options);
            var markdownPath = Path.Combine(outputPath, "Project Notebook", "Meeting Notes", "Planning Page.md");

            result.Success.Should().BeTrue();
            result.Warnings.Should().ContainSingle().Which.Should().Contain("Markdown linting failed for 'Planning Page': lint failed");
            linter.Calls.Should().Be(1);
            File.ReadAllText(markdownPath).Should().StartWith("asset:");
        }
        finally
        {
            if (Directory.Exists(outputPath)) Directory.Delete(outputPath, true);
        }
    }

    [Fact]
    public async Task ExportSelectedAsync_WithSelectedParentPage_ExportsParentAndSubpagesInNestedFolders()
    {
        var outputPath = CreateTempExportPath();
        var service = new ExportService(new FakeOneNoteExportSource(), new RecordingMarkdownConverter(), new PassthroughMarkdownLintService());
        var notebooks = CreateParentSubpageHierarchy(parentSelected: true, childSelected: false);
        var options = new ExportOptions { OutputPath = outputPath, ApplyLinting = false, Overwrite = true };

        try
        {
            var result = await service.ExportSelectedAsync(notebooks, options);

            result.Success.Should().BeTrue();
            File.Exists(Path.Combine(outputPath, "Notebook", "Section", "Parent Page.md")).Should().BeTrue();
            File.Exists(Path.Combine(outputPath, "Notebook", "Section", "Parent Page", "Child Page.md")).Should().BeTrue();
            File.Exists(Path.Combine(outputPath, "Notebook", "Section", "Parent Page", "Child Page", "Grandchild Page.md")).Should().BeTrue();
        }
        finally
        {
            if (Directory.Exists(outputPath)) Directory.Delete(outputPath, true);
        }
    }

    [Fact]
    public async Task ExportSelectedAsync_WithOnlySelectedSubpage_CreatesParentFolderWithoutParentMarkdown()
    {
        var outputPath = CreateTempExportPath();
        var service = new ExportService(new FakeOneNoteExportSource(), new RecordingMarkdownConverter(), new PassthroughMarkdownLintService());
        var notebooks = CreateParentSubpageHierarchy(parentSelected: false, childSelected: true);
        var options = new ExportOptions { OutputPath = outputPath, ApplyLinting = false, Overwrite = true };

        try
        {
            var result = await service.ExportSelectedAsync(notebooks, options);

            result.Success.Should().BeTrue();
            File.Exists(Path.Combine(outputPath, "Notebook", "Section", "Parent Page.md")).Should().BeFalse();
            Directory.Exists(Path.Combine(outputPath, "Notebook", "Section", "Parent Page")).Should().BeTrue();
            File.Exists(Path.Combine(outputPath, "Notebook", "Section", "Parent Page", "Child Page.md")).Should().BeTrue();
        }
        finally
        {
            if (Directory.Exists(outputPath)) Directory.Delete(outputPath, true);
        }
    }

    private static List<OneNoteItem> CreateSelectedHierarchy(DateTimeOffset? created = null, DateTimeOffset? modified = null)
    {
        return new List<OneNoteItem>
        {
            new OneNoteItem
            {
                Id = "notebook-id",
                Name = "Project Notebook",
                Type = OneNoteItemType.Notebook,
                IsSelected = true,
                Children =
                {
                    new OneNoteItem
                    {
                        Id = "section-id",
                        Name = "Meeting Notes",
                        Type = OneNoteItemType.Section,
                        Children =
                        {
                            new OneNoteItem
                            {
                                Id = "page-id",
                                Name = "Planning Page",
                                Type = OneNoteItemType.Page,
                                CreatedTime = created,
                                LastModifiedTime = modified
                            }
                        }
                    }
                }
            }
        };
    }

    private static List<OneNoteItem> CreateParentSubpageHierarchy(bool parentSelected, bool childSelected)
    {
        return new List<OneNoteItem>
        {
            new OneNoteItem
            {
                Id = "notebook-id",
                Name = "Notebook",
                Type = OneNoteItemType.Notebook,
                Children =
                {
                    new OneNoteItem
                    {
                        Id = "section-id",
                        Name = "Section",
                        Type = OneNoteItemType.Section,
                        Children =
                        {
                            new OneNoteItem
                            {
                                Id = "parent-page-id",
                                Name = "Parent Page",
                                Type = OneNoteItemType.Page,
                                IsSelected = parentSelected,
                                Children =
                                {
                                    new OneNoteItem
                                    {
                                        Id = "child-page-id",
                                        Name = "Child Page",
                                        Type = OneNoteItemType.Page,
                                        IsSelected = childSelected,
                                        Children =
                                        {
                                            new OneNoteItem
                                            {
                                                Id = "grandchild-page-id",
                                                Name = "Grandchild Page",
                                                Type = OneNoteItemType.Page
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        };
    }

    private static string CreateTempExportPath()
    {
        return Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
    }

    private sealed class FakeOneNoteExportSource : IOneNoteExportSource
    {
        public List<OneNoteItem> GetNotebookHierarchy()
        {
            return CreateSelectedHierarchy();
        }

        public string GetPageContent(string pageId)
        {
            return "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" />";
        }

        public string? GetBinaryPageContent(string pageId, string callbackId)
        {
            return null;
        }
    }

    private sealed class RecordingMarkdownConverter : IMarkdownContentConverter
    {
        public List<ConverterCall> Calls { get; } = new();

        public string Convert(string pageXml, string assetsFolder, string relativeAssetsPath, BinaryContentFetcher? binaryContentFetcher = null, string? pagePrefix = null)
        {
            Calls.Add(new ConverterCall(assetsFolder, relativeAssetsPath));
            return $"asset:{relativeAssetsPath}";
        }
    }

    private sealed class AssetWritingMarkdownConverter : IMarkdownContentConverter
    {
        public string Convert(string pageXml, string assetsFolder, string relativeAssetsPath, BinaryContentFetcher? binaryContentFetcher = null, string? pagePrefix = null)
        {
            Directory.CreateDirectory(assetsFolder);
            File.WriteAllBytes(Path.Combine(assetsFolder, "image_0001.png"), new byte[] { 1, 2, 3 });
            return $"![image]({relativeAssetsPath}/image_0001.png)";
        }
    }

    private sealed record ConverterCall(string AssetsFolder, string RelativeAssetsPath);

    private sealed class PassthroughMarkdownLintService : IMarkdownLintService
    {
        public bool IsAvailable => true;
        public string UnavailableReason => string.Empty;

        public LintResult LintContent(string markdown, string? configPath = null)
        {
            return new LintResult { Success = true, Content = markdown };
        }
    }

    private sealed class RecordingMarkdownLintService : IMarkdownLintService
    {
        public bool IsAvailable => true;
        public string UnavailableReason => string.Empty;
        public List<string> Inputs { get; } = new();
        public List<string?> ConfigPaths { get; } = new();

        public LintResult LintContent(string markdown, string? configPath = null)
        {
            Inputs.Add(markdown);
            ConfigPaths.Add(configPath);
            return new LintResult { Success = true, Content = markdown };
        }
    }

    private sealed class StubMarkdownLintService : IMarkdownLintService
    {
        public bool IsAvailable { get; init; }
        public string UnavailableReason { get; init; } = string.Empty;
        public LintResult Result { get; init; } = new() { Success = true };
        public int Calls { get; private set; }

        public LintResult LintContent(string markdown, string? configPath = null)
        {
            Calls++;
            return Result;
        }
    }

    private sealed class RecordingFileTimestampService : IFileTimestampService
    {
        public List<TimestampCall> Calls { get; } = new();

        public void ApplyTimestamps(string filePath, DateTimeOffset? created, DateTimeOffset? modified)
        {
            Calls.Add(new TimestampCall(filePath, created, modified));
        }
    }

    private sealed class ThrowingFileTimestampService : IFileTimestampService
    {
        public void ApplyTimestamps(string filePath, DateTimeOffset? created, DateTimeOffset? modified)
        {
            throw new IOException("timestamp failed");
        }
    }

    private sealed class PassthroughYamlFrontMatterService : IYamlFrontMatterService
    {
        public string AddFrontMatter(string markdown, OneNoteItem page)
        {
            return markdown;
        }
    }

    private sealed class RecordingYamlFrontMatterService : IYamlFrontMatterService
    {
        public int Calls { get; private set; }

        public string AddFrontMatter(string markdown, OneNoteItem page)
        {
            Calls++;
            return $"---\nmetadata: true\n---\n\n{markdown}";
        }
    }

    private sealed record TimestampCall(string FilePath, DateTimeOffset? Created, DateTimeOffset? Modified);
}
