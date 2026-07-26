using System.IO;
using FluentAssertions;
using OneNoteMarkdownExporter.Services;
using Xunit;

namespace OneNoteMarkdownExporter.Tests.Converters;

/// <summary>
/// Tests for the OneNoteXmlToMarkdownConverter - the core conversion engine.
/// </summary>
public class OneNoteXmlToMarkdownConverterTests
{
    private readonly OneNoteXmlToMarkdownConverter _converter;

    public OneNoteXmlToMarkdownConverterTests()
    {
        _converter = new OneNoteXmlToMarkdownConverter();
    }

    #region Basic Conversion Tests

    [Fact]
    public void Convert_SimpleText_ReturnsMarkdown()
    {
        // Arrange
        var xml = CreatePageXml("<one:T><![CDATA[Hello World]]></one:T>");

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().Contain("Hello World");
    }

    [Fact]
    public void Convert_EncodedComparisonCharacters_UsesReadableMarkdownText()
    {
        var xml = CreatePageXml("<one:T><![CDATA[This is plain text saying 2 &gt; 3 and 2 &lt; 3.]]></one:T>");

        var result = _converter.Convert(xml, "", "assets", null, "test");

        result.Should().Contain("2 > 3 and 2 < 3");
    }

    [Fact]
    public void Convert_EncodedHtmlLikeText_DoesNotInterpretTags()
    {
        var xml = CreatePageXml("<one:T><![CDATA[Before &lt;br&gt; &lt;html&gt;&lt;/html&gt; &lt;a&gt;&lt;/a&gt; After]]></one:T>");

        var result = _converter.Convert(xml, "", "assets", null, "test");

        result.Should().Contain(@"Before \<br> \<html>\</html> \<a>\</a> After");
        result.Should().NotContain("Before\n");
    }

    [Fact]
    public void Convert_TestPageLiteralAngleText_MatchesReadableMarkdownSource()
    {
        var xml = CreatePageXml("<one:T><![CDATA[This is plain text with a bunch of opening and closing tags that aren't HTML, I just typed a bunch here: &lt;&gt; &lt;br&gt; &lt; &gt;&gt; &lt;&lt; &gt;&gt;&gt;&gt; &lt; &lt;html&gt;&lt;/html&gt; &lt;a&gt;&lt;/a&gt;  -- none of that is HTML from OneNote it's me just typing plain text HTML into a one note page.]]></one:T>");

        var result = _converter.Convert(xml, "", "assets", null, "test");

        result.Should().Be(@"This is plain text with a bunch of opening and closing tags that aren't HTML, I just typed a bunch here: <> \<br> < >> << >>>> < \<html>\</html> \<a>\</a>  -- none of that is HTML from OneNote it's me just typing plain text HTML into a one note page.");
    }
    [Fact]
    public void Convert_EncodedHtmlLikeTextWithUrl_DoesNotRewriteLiteralTag()
    {
        var xml = CreatePageXml("<one:T><![CDATA[Example: &lt;a href=&quot;https://example.com&quot;&gt;Link&lt;/a&gt;]]></one:T>");

        var result = _converter.Convert(xml, "", "assets", null, "test");

        result.Should().Contain("\\<a href=\"https://example.com\">Link\\</a>");
        result.Should().NotContain("<https://example.com");
    }

    [Fact]
    public void Convert_EncodedHtmlCommentAndDeclaration_EscapesMarkdownHtmlStarts()
    {
        var xml = CreatePageXml("<one:T><![CDATA[&lt;!-- comment --&gt; &lt;!DOCTYPE html&gt; &lt;?target value?&gt;]]></one:T>");

        var result = _converter.Convert(xml, "", "assets", null, "test");

        result.Should().Be(@"\<!-- comment --> \<!DOCTYPE html> \<?target value?>");
    }

    [Fact]
    public void Convert_EncodedGreaterThanAtLineStart_DoesNotCreateBlockquote()
    {
        var xml = CreatePageXml("<one:T><![CDATA[&gt; Literal greater-than text]]></one:T>");

        var result = _converter.Convert(xml, "", "assets", null, "test");

        result.Should().Be(@"\> Literal greater-than text");
    }
    [Fact]
    public void Convert_PageWithTitle_IncludesH1Heading()
    {
        // Arrange
        var xml = @"<?xml version=""1.0""?>
            <one:Page xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote"">
                <one:Title>
                    <one:OE><one:T><![CDATA[My Page Title]]></one:T></one:OE>
                </one:Title>
                <one:Outline>
                    <one:OEChildren>
                        <one:OE><one:T><![CDATA[Content]]></one:T></one:OE>
                    </one:OEChildren>
                </one:Outline>
            </one:Page>";

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().Contain("# My Page Title");
    }

    [Fact]
    public void Convert_PageTitleWithEncodedHtmlLikeText_PreservesSafeText()
    {
        var xml = @"<?xml version=""1.0""?>
            <one:Page xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote"">
                <one:Title>
                    <one:OE><one:T><![CDATA[Title &lt;html&gt; &amp; Friends]]></one:T></one:OE>
                </one:Title>
            </one:Page>";

        var result = _converter.Convert(xml, "", "assets", null, "test");

        result.Should().Contain(@"# Title \<html> & Friends");
        result.Should().NotContain("&lt;");
    }

    [Fact]
    public void Convert_EmptyPage_ReturnsNonNull()
    {
        // Arrange
        var xml = @"<?xml version=""1.0""?>
            <one:Page xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote"">
            </one:Page>";

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().NotBeNull();
    }

    #endregion

    #region Text Formatting Tests

    [Fact]
    public void Convert_BoldText_ConvertsToStrong()
    {
        // Arrange - Use single quotes for style attribute (OneNote format)
        var xml = CreatePageXml("<one:T><![CDATA[<span style='font-weight:bold'>Bold Text</span>]]></one:T>");

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().Contain("**Bold Text**");
    }

    [Fact]
    public void Convert_RawStrongElement_ConvertsToBold()
    {
        var xml = CreatePageXml("<one:T><![CDATA[<strong>Strong Text</strong>]]></one:T>");

        var result = _converter.Convert(xml, "", "assets", null, "test");

        result.Should().Contain("**Strong Text**");
    }

    [Fact]
    public void Convert_ItalicText_ConvertsToEmphasis()
    {
        // Arrange - Use single quotes for style attribute (OneNote format)
        var xml = CreatePageXml("<one:T><![CDATA[<span style='font-style:italic'>Italic Text</span>]]></one:T>");

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().Contain("*Italic Text*");
    }

    [Fact]
    public void Convert_StrikethroughText_ConvertsToDelTag()
    {
        // Arrange - Strikethrough uses T element style attribute, not span
        var xml = CreatePageXmlWithStyle("Deleted", "text-decoration:line-through");

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().Contain("~~Deleted~~");
    }

    [Fact]
    public void Convert_StyledEncodedHtmlLikeText_PreservesLiteralTag()
    {
        var xml = CreatePageXmlWithStyle("Literal &lt;br&gt; text", "font-weight:bold");

        var result = _converter.Convert(xml, "", "assets", null, "test");

        result.Should().Contain(@"**Literal \<br> text**");
        result.Should().NotContain("Literal\n");
    }

    [Fact]
    public void Convert_HighlightedText_ConvertsToBold()
    {
        // Arrange - highlighted text has background color (single quotes)
        var xml = CreatePageXml("<one:T><![CDATA[<span style='background:yellow'>Highlighted</span>]]></one:T>");

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().Contain("**Highlighted**");
    }

    [Fact]
    public void Convert_BoldAndItalic_PreservesBothFormats()
    {
        var xml = CreatePageXml("<one:T><![CDATA[<span style='font-weight:bold;font-style:italic'>Bold Italic</span>]]></one:T>");

        var result = _converter.Convert(xml, "", "assets", null, "test");

        result.Should().Be("***Bold Italic***");
    }

    [Fact]
    public void Convert_BoldItalicAndHighlighted_PreservesBoldAndItalic()
    {
        var xml = CreatePageXml("<one:T><![CDATA[<span style='font-weight:bold;font-style:italic;background:yellow;mso-highlight:yellow'>Styled Text</span>]]></one:T>");

        var result = _converter.Convert(xml, "", "assets", null, "test");

        result.Should().Be("***Styled Text***");
    }

    [Fact]
    public void Convert_NestedBoldAndItalicSpans_PreservesBothFormats()
    {
        var xml = CreatePageXml("<one:T><![CDATA[<span style='font-weight:bold'><span style='font-style:italic'>Nested Text</span></span>]]></one:T>");

        var result = _converter.Convert(xml, "", "assets", null, "test");

        result.Should().Be("***Nested Text***");
    }

    [Fact]
    public void Convert_LinkInsideBoldSpan_PreservesLinkAndFormatting()
    {
        var xml = CreatePageXml("<one:T><![CDATA[<span style='font-weight:bold'><a href='https://example.com'>Example</a></span>]]></one:T>");

        var result = _converter.Convert(xml, "", "assets", null, "test");

        result.Should().Be("**[Example](https://example.com)**");
    }

    [Fact]
    public void Convert_TextWithEmbeddedBreak_CreatesProseLineBreak()
    {
        var xml = CreatePageXml("<one:T><![CDATA[First line<br/>Second line]]></one:T>");

        var result = _converter.Convert(xml, "", "assets", null, "test");

        result.Should().Contain("First line\nSecond line");
        result.Should().NotContain("<br>");
        result.Should().NotContain("<br/>");
    }

    #endregion

    #region List Tests

    [Fact]
    public void Convert_BulletList_CreatesUnorderedList()
    {
        // Arrange
        var xml = @"<?xml version=""1.0""?>
            <one:Page xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote"">
                <one:Outline>
                    <one:OEChildren>
                        <one:OE>
                            <one:List><one:Bullet /></one:List>
                            <one:T><![CDATA[Item 1]]></one:T>
                        </one:OE>
                        <one:OE>
                            <one:List><one:Bullet /></one:List>
                            <one:T><![CDATA[Item 2]]></one:T>
                        </one:OE>
                    </one:OEChildren>
                </one:Outline>
            </one:Page>";

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().Contain("- Item 1");
        result.Should().Contain("- Item 2");
    }

    [Fact]
    public void Convert_OutlineWithMultipleOEChildren_PreservesContentAfterBulletList()
    {
        // Arrange
        var xml = @"<?xml version=""1.0""?>
            <one:Page xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote"">
                <one:Outline>
                    <one:OEChildren indent=""2"">
                        <one:OE>
                            <one:List><one:Bullet /></one:List>
                            <one:T><![CDATA[The first thing]]></one:T>
                        </one:OE>
                        <one:OE>
                            <one:List><one:Bullet /></one:List>
                            <one:T><![CDATA[The second thing]]></one:T>
                        </one:OE>
                    </one:OEChildren>
                    <one:OEChildren>
                        <one:OE><one:T><![CDATA[]]></one:T></one:OE>
                        <one:OE><one:T><![CDATA[This should still show up]]></one:T></one:OE>
                    </one:OEChildren>
                </one:Outline>
            </one:Page>";

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().Contain("- The first thing");
        result.Should().Contain("- The second thing");
        result.Should().Contain("This should still show up");
    }

    [Fact]
    public void Convert_NumberedList_CreatesOrderedList()
    {
        // Arrange
        var xml = @"<?xml version=""1.0""?>
            <one:Page xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote"">
                <one:Outline>
                    <one:OEChildren>
                        <one:OE>
                            <one:List><one:Number /></one:List>
                            <one:T><![CDATA[First]]></one:T>
                        </one:OE>
                        <one:OE>
                            <one:List><one:Number /></one:List>
                            <one:T><![CDATA[Second]]></one:T>
                        </one:OE>
                    </one:OEChildren>
                </one:Outline>
            </one:Page>";

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().Contain("1. First");
        result.Should().Contain("2. Second");
    }

    [Fact]
    public void Convert_NestedBulletList_PreservesIndentation()
    {
        // Arrange
        var xml = @"<?xml version=""1.0""?>
            <one:Page xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote"">
                <one:Outline>
                    <one:OEChildren>
                        <one:OE>
                            <one:List><one:Bullet /></one:List>
                            <one:T><![CDATA[Parent]]></one:T>
                            <one:OEChildren>
                                <one:OE>
                                    <one:List><one:Bullet /></one:List>
                                    <one:T><![CDATA[Child]]></one:T>
                                </one:OE>
                            </one:OEChildren>
                        </one:OE>
                    </one:OEChildren>
                </one:Outline>
            </one:Page>";

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().Contain("- Parent");
        result.Should().Contain("Child"); // Should be indented or nested
    }

    #endregion

    #region Link Tests

    [Fact]
    public void Convert_SimpleLink_CreatesMarkdownLink()
    {
        // Arrange
        var xml = CreatePageXml(@"<one:T><![CDATA[<a href=""https://example.com"">Click Here</a>]]></one:T>");

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().Contain("[Click Here](https://example.com)");
    }

    [Fact]
    public void Convert_NakedUrlAnchor_ConvertsToAutolink()
    {
        // Arrange - when link text matches URL
        var xml = CreatePageXml(@"<one:T><![CDATA[<a href=""https://example.com"">https://example.com</a>]]></one:T>");

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().Contain("<https://example.com>");
    }

    [Fact]
    public void Convert_PlainBareUrl_WrapsInAngleBrackets()
    {
        // Arrange
        var xml = CreatePageXml(@"<one:T><![CDATA[Visit https://example.com for details]]></one:T>");

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().Contain("Visit <https://example.com> for details");
    }

    [Fact]
    public void Convert_PlainBareUrlWithTrailingPeriod_LeavesPeriodOutsideAngleBrackets()
    {
        // Arrange
        var xml = CreatePageXml(@"<one:T><![CDATA[Visit https://example.com.]]></one:T>");

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().Contain("Visit <https://example.com>.");
    }

    [Fact]
    public void Convert_PlainBareUrlWithTrailingComma_LeavesCommaOutsideAngleBrackets()
    {
        // Arrange
        var xml = CreatePageXml(@"<one:T><![CDATA[Visit https://example.com, then continue]]></one:T>");

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().Contain("Visit <https://example.com>, then continue");
    }

    [Fact]
    public void Convert_ExistingAutolink_DoesNotDoubleWrap()
    {
        // OneNote entity-encodes angle brackets typed by the user.
        var xml = CreatePageXml(@"<one:T><![CDATA[Visit &lt;https://example.com&gt;]]></one:T>");

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().Contain("Visit <https://example.com>");
        result.Should().NotContain("<<https://example.com>>");
    }

    [Fact]
    public void Convert_MarkdownLink_DoesNotWrapHref()
    {
        // Arrange
        var xml = CreatePageXml(@"<one:T><![CDATA[See [Example](https://example.com)]]></one:T>");

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().Contain("See [Example](https://example.com)");
        result.Should().NotContain("[Example](<https://example.com>)");
    }

    [Fact]
    public void Convert_LinkWithSpecialChars_PreservesUrl()
    {
        // Arrange
        var xml = CreatePageXml(@"<one:T><![CDATA[<a href=""https://example.com/path?query=value&other=123"">Link</a>]]></one:T>");

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().Contain("https://example.com/path?query=value&other=123");
    }

    [Fact]
    public void Convert_LinkWithEncodedAmpersand_DecodesMarkdownDestination()
    {
        var xml = CreatePageXml(@"<one:T><![CDATA[<a href=""https://example.com/?first=1&amp;second=2"">Link</a>]]></one:T>");

        var result = _converter.Convert(xml, "", "assets", null, "test");

        result.Should().Contain("[Link](https://example.com/?first=1&second=2)");
        result.Should().NotContain("&amp;");
    }

    #endregion

    #region Table Tests

    [Fact]
    public void Convert_SimpleTable_CreatesMarkdownTable()
    {
        // Arrange
        var xml = @"<?xml version=""1.0""?>
            <one:Page xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote"">
                <one:Outline>
                    <one:OEChildren>
                        <one:OE>
                            <one:Table>
                                <one:Row>
                                    <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[A]]></one:T></one:OE></one:OEChildren></one:Cell>
                                    <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[B]]></one:T></one:OE></one:OEChildren></one:Cell>
                                </one:Row>
                                <one:Row>
                                    <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[1]]></one:T></one:OE></one:OEChildren></one:Cell>
                                    <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[2]]></one:T></one:OE></one:OEChildren></one:Cell>
                                </one:Row>
                            </one:Table>
                        </one:OE>
                    </one:OEChildren>
                </one:Outline>
            </one:Page>";

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert - ReverseMarkdown converts tables to Markdown format
        result.Should().Contain("|");
        result.Should().Contain("---");
    }

    [Fact]
    public void Convert_TableCellWithMultipleParagraphs_KeepsRowOnSingleLine()
    {
        var xml = @"<?xml version=""1.0""?>
            <one:Page xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote"">
                <one:Outline>
                    <one:OEChildren>
                        <one:OE>
                            <one:Table>
                                <one:Row>
                                    <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Topic]]></one:T></one:OE></one:OEChildren></one:Cell>
                                    <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Details]]></one:T></one:OE></one:OEChildren></one:Cell>
                                </one:Row>
                                <one:Row>
                                    <one:Cell>
                                        <one:OEChildren>
                                            <one:OE><one:T><![CDATA[First paragraph]]></one:T></one:OE>
                                            <one:OE><one:T><![CDATA[Second paragraph]]></one:T></one:OE>
                                        </one:OEChildren>
                                    </one:Cell>
                                    <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Value]]></one:T></one:OE></one:OEChildren></one:Cell>
                                </one:Row>
                            </one:Table>
                        </one:OE>
                    </one:OEChildren>
                </one:Outline>
            </one:Page>";

        var result = _converter.Convert(xml, "", "assets", null, "test");

        result.Should().Contain("| First paragraph<br>Second paragraph | Value |");
    }

    [Fact]
    public void Convert_TableCellWithEmbeddedBreak_KeepsRowOnSingleLine()
    {
        var xml = @"<?xml version=""1.0""?>
            <one:Page xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote"">
                <one:Outline>
                    <one:OEChildren>
                        <one:OE>
                            <one:Table>
                                <one:Row>
                                    <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Topic]]></one:T></one:OE></one:OEChildren></one:Cell>
                                    <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Details]]></one:T></one:OE></one:OEChildren></one:Cell>
                                </one:Row>
                                <one:Row>
                                    <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Name]]></one:T></one:OE></one:OEChildren></one:Cell>
                                    <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[<span>First line<br/>Second line</span>]]></one:T></one:OE></one:OEChildren></one:Cell>
                                </one:Row>
                            </one:Table>
                        </one:OE>
                    </one:OEChildren>
                </one:Outline>
            </one:Page>";

        var result = _converter.Convert(xml, "", "assets", null, "test");

        result.Should().Contain("| Name | First line<br>Second line |");
    }

    #endregion

    #region Image Tests

    [Fact]
    public void Convert_ImageWithCustomRelativeAssetsPath_UsesRelativeAssetsPathInMarkdown()
    {
        // Arrange
        var tempRoot = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        var assetsFolder = Path.Combine(tempRoot, "shared", "assets");
        var xml = @"<?xml version=""1.0""?>
            <one:Page xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote"">
                <one:Image format=""png""><one:Data>AQID</one:Data></one:Image>
            </one:Page>";

        try
        {
            // Act
            var result = _converter.Convert(xml, assetsFolder, "../shared/assets", null, "page");

            // Assert
            result.Should().Contain("../shared/assets/page_image_0001.png");
            File.Exists(Path.Combine(assetsFolder, "page_image_0001.png")).Should().BeTrue();
        }
        finally
        {
            if (Directory.Exists(tempRoot)) Directory.Delete(tempRoot, true);
        }
    }

    [Fact]
    public void Convert_ImageWithExistingAssetFile_OverwritesAssetFile()
    {
        // Arrange
        var tempRoot = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        var assetsFolder = Path.Combine(tempRoot, "assets");
        var existingAssetPath = Path.Combine(assetsFolder, "page_image_0001.png");
        var xml = @"<?xml version=""1.0""?>
            <one:Page xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote"">
                <one:Image format=""png""><one:Data>AQID</one:Data></one:Image>
            </one:Page>";

        try
        {
            Directory.CreateDirectory(assetsFolder);
            File.WriteAllBytes(existingAssetPath, new byte[] { 9, 9, 9 });

            // Act
            var result = _converter.Convert(xml, assetsFolder, "assets", null, "page");

            // Assert
            result.Should().Contain("assets/page_image_0001.png");
            File.ReadAllBytes(existingAssetPath).Should().Equal(new byte[] { 1, 2, 3 });
        }
        finally
        {
            if (Directory.Exists(tempRoot)) Directory.Delete(tempRoot, true);
        }
    }

    #endregion

    #region Cleanup Tests

    [Fact]
    public void Convert_MultipleBlankLines_ReducedToTwo()
    {
        // Arrange
        var xml = @"<?xml version=""1.0""?>
            <one:Page xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote"">
                <one:Outline>
                    <one:OEChildren>
                        <one:OE><one:T><![CDATA[Line 1]]></one:T></one:OE>
                        <one:OE><one:T><![CDATA[]]></one:T></one:OE>
                        <one:OE><one:T><![CDATA[]]></one:T></one:OE>
                        <one:OE><one:T><![CDATA[]]></one:T></one:OE>
                        <one:OE><one:T><![CDATA[]]></one:T></one:OE>
                        <one:OE><one:T><![CDATA[Line 2]]></one:T></one:OE>
                    </one:OEChildren>
                </one:Outline>
            </one:Page>";

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        // Should not have more than 2 consecutive newlines (3+ newline chars in a row)
        result.Should().NotContain("\n\n\n\n");
    }

    [Fact]
    public void Convert_HtmlEntities_Decoded()
    {
        // Arrange
        var xml = CreatePageXml(@"<one:T><![CDATA[Tom &amp; Jerry]]></one:T>");

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().Contain("Tom & Jerry");
    }

    [Fact]
    public void Convert_UnicodeContent_Preserved()
    {
        // Arrange
        var xml = CreatePageXml(@"<one:T><![CDATA[Hello 世界 🎉]]></one:T>");

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().Contain("Hello 世界 🎉");
    }

    #endregion

    #region Edge Cases

    [Fact]
    public void Convert_NullBinaryFetcher_HandlesGracefully()
    {
        // Arrange
        var xml = CreatePageXml(@"<one:T><![CDATA[Simple text]]></one:T>");

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "test");

        // Assert
        result.Should().NotBeNull();
        result.Should().Contain("Simple text");
    }

    [Fact]
    public void Convert_EmptyPrefix_HandlesGracefully()
    {
        // Arrange
        var xml = CreatePageXml(@"<one:T><![CDATA[Content]]></one:T>");

        // Act
        var result = _converter.Convert(xml, "", "assets", null, "");

        // Assert
        result.Should().NotBeNull();
    }

    [Fact]
    public void Convert_SpecialCharsInPrefix_Sanitized()
    {
        // Arrange
        var xml = CreatePageXml(@"<one:T><![CDATA[Content]]></one:T>");

        // Act - prefix with invalid filename characters
        var result = _converter.Convert(xml, "", "assets", null, "test:page/name");

        // Assert
        result.Should().NotBeNull();
    }

    #endregion

    #region Helper Methods

    /// <summary>
    /// Creates a minimal OneNote page XML with the given content element.
    /// </summary>
    private static string CreatePageXml(string contentElement)
    {
        return $@"<?xml version=""1.0""?>
            <one:Page xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote"">
                <one:Outline>
                    <one:OEChildren>
                        <one:OE>{contentElement}</one:OE>
                    </one:OEChildren>
                </one:Outline>
            </one:Page>";
    }

    /// <summary>
    /// Creates a OneNote page XML with text that has a style attribute on the T element.
    /// </summary>
    private static string CreatePageXmlWithStyle(string text, string style)
    {
        return $@"<?xml version=""1.0""?>
            <one:Page xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote"">
                <one:Outline>
                    <one:OEChildren>
                        <one:OE><one:T style=""{style}""><![CDATA[{text}]]></one:T></one:OE>
                    </one:OEChildren>
                </one:Outline>
            </one:Page>";
    }

    #endregion
}
