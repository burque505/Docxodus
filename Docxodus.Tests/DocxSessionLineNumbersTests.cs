#nullable enable

using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Word's line numbers: a section's <c>w:lnNumType</c> through <see cref="PageSetupOp.LineNumbering"/>
/// and <see cref="SectionInfo.LineNumbering"/>, a paragraph's <c>w:suppressLineNumbers</c> through
/// <see cref="ParagraphFormatOp.SuppressLineNumbers"/>, and the HTML attributes the paginated view
/// draws them from. Test IDs use the LN0xx range.
/// </summary>
public class DocxSessionLineNumbersTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static string FirstParagraph(DocxSession session) =>
        session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;

    private static XElement? LnNumType(DocxSession session)
    {
        using var stream = new MemoryStream(session.Save());
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.GetXDocument().Root!.Descendants(W + "sectPr").LastOrDefault()?.Element(W + "lnNumType");
    }

    [Fact]
    public void LN001_LineNumbering_IsWrittenAsWordWritesIt_AndReadsBack()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstParagraph(session);

        Assert.True(session.SetPageSetup(anchor, new PageSetupOp { LineNumbering = new LineNumbering() }).Success);
        // Word: Start at 1, Auto distance and Restart each page are all left out.
        var ln = LnNumType(session)!;
        Assert.Equal(new[] { "countBy" }, ln.Attributes().Select(a => a.Name.LocalName));
        Assert.Equal("1", (string?)ln.Attribute(W + "countBy"));
        Assert.Equal(new LineNumbering(), session.GetSectionInfo(anchor)!.LineNumbering);

        var wanted = new LineNumbering { CountBy = 2, Start = 5, DistanceTwips = 360, Restart = LineNumberRestart.NewSection };
        Assert.True(session.SetPageSetup(anchor, new PageSetupOp { LineNumbering = wanted }).Success);
        ln = LnNumType(session)!;
        // w:start is one less than the first number shown (Word writes Start at 5 as 4).
        Assert.Equal("4", (string?)ln.Attribute(W + "start"));
        Assert.Equal("360", (string?)ln.Attribute(W + "distance"));
        Assert.Equal("newSection", (string?)ln.Attribute(W + "restart"));
        Assert.Equal(wanted, session.GetSectionInfo(anchor)!.LineNumbering);
        PackageValidation.AssertValid(session.Save(), DocumentFormat.OpenXml.FileFormatVersions.Office2007);
    }

    [Fact]
    public void LN002_CountByZero_RemovesIt_TheSameValue_IsNoEdit_AndUndoRestores()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstParagraph(session);
        var on = new LineNumbering { Restart = LineNumberRestart.Continuous };
        Assert.True(session.SetPageSetup(anchor, new PageSetupOp { LineNumbering = on }).Success);
        Assert.Equal("continuous", (string?)LnNumType(session)!.Attribute(W + "restart"));

        long version = session.Version;
        Assert.True(session.SetPageSetup(anchor, new PageSetupOp { LineNumbering = on }).Success);
        Assert.Equal(version, session.Version);

        Assert.True(session.SetPageSetup(anchor, new PageSetupOp { LineNumbering = new LineNumbering { CountBy = 0 } }).Success);
        Assert.Null(LnNumType(session));
        Assert.Null(session.GetSectionInfo(anchor)!.LineNumbering);

        Assert.True(session.Undo());
        Assert.Equal(on, session.GetSectionInfo(anchor)!.LineNumbering);
    }

    [Theory]
    [InlineData(-1, 1, null)]
    [InlineData(101, 1, null)]
    [InlineData(1, 0, null)]
    [InlineData(1, 1, -10)]
    public void LN003_InvalidLineNumbering_IsRefused(int countBy, int start, int? distance)
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstParagraph(session);
        var result = session.SetPageSetup(anchor, new PageSetupOp
        {
            LineNumbering = new LineNumbering { CountBy = countBy, Start = start, DistanceTwips = distance },
        });
        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.InvalidPageSetup, result.Error!.Code);
        Assert.Null(LnNumType(session));
    }

    [Fact]
    public void LN004_SuppressLineNumbers_TogglesTheParagraphProperty_AndReadsBack()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstParagraph(session);
        var on = session.SetParagraphFormat(anchor, new ParagraphFormatOp { SuppressLineNumbers = true });
        Assert.True(on.Success, on.Error?.Message);
        anchor = on.Modified[0].Id;
        Assert.True(session.GetFormatting(anchor)!.DirectParagraph.SuppressLineNumbers);

        var off = session.SetParagraphFormat(anchor, new ParagraphFormatOp { SuppressLineNumbers = false });
        Assert.True(off.Success, off.Error?.Message);
        Assert.False(session.GetFormatting(off.Modified[0].Id)!.EffectiveParagraph.SuppressLineNumbers);
    }

    [Fact]
    public void LN005_TheHtmlCarriesTheSectionsLineNumbering_AndASuppressedParagraph()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstParagraph(session);
        Assert.True(session.SetPageSetup(anchor, new PageSetupOp
        {
            LineNumbering = new LineNumbering { CountBy = 5, Start = 3, DistanceTwips = 240, Restart = LineNumberRestart.Continuous },
        }).Success);
        Assert.True(session.SetParagraphFormat(anchor, new ParagraphFormatOp { SuppressLineNumbers = true }).Success);

        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("ln.docx", session.Save()), new WmlToHtmlConverterSettings { RenderPagination = PaginationMode.Paginated })
            .ToString(SaveOptions.DisableFormatting);
        Assert.Contains("data-line-numbers=\"5\"", html);
        Assert.Contains("data-line-number-start=\"3\"", html);
        Assert.Contains("data-line-number-distance=\"12.0\"", html);
        Assert.Contains("data-line-number-restart=\"continuous\"", html);
        Assert.Contains("data-suppress-line-numbers=\"true\"", html);
    }

    [Fact]
    public void LN006_Json_ParsesAndSerializesLineNumbering()
    {
        var op = DocxSessionJson.ParsePageSetupOp(
            "{\"lineNumbering\":{\"countBy\":2,\"start\":10,\"restart\":\"newSection\"}}");
        Assert.Equal(new LineNumbering { CountBy = 2, Start = 10, Restart = LineNumberRestart.NewSection }, op.LineNumbering);
        Assert.Null(DocxSessionJson.ParsePageSetupOp("{\"marginLeftTwips\":720}").LineNumbering);

        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstParagraph(session);
        Assert.True(session.SetPageSetup(anchor, new PageSetupOp { MarginLeftTwips = 1440 }).Success);
        using (var none = JsonDocument.Parse(DocxSessionJson.SerializeSectionInfoOrNull(session.GetSectionInfo(anchor))))
            Assert.False(none.RootElement.TryGetProperty("lineNumbering", out _));
        Assert.True(session.SetPageSetup(anchor, op).Success);
        using var json = JsonDocument.Parse(DocxSessionJson.SerializeSectionInfoOrNull(session.GetSectionInfo(anchor)));
        var ln = json.RootElement.GetProperty("lineNumbering");
        Assert.Equal(2, ln.GetProperty("countBy").GetInt32());
        Assert.Equal(10, ln.GetProperty("start").GetInt32());
        Assert.Equal("newSection", ln.GetProperty("restart").GetString());
        Assert.False(ln.TryGetProperty("distanceTwips", out _));
    }
}
