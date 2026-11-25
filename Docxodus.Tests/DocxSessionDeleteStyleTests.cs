using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Deleting a custom style with <see cref="DocxSession.DeleteStyle"/> (Word's Styles pane) and the
/// <see cref="StyleInfo.InUse"/> flag. Test IDs continue the SD0xx range.
/// </summary>
public class DocxSessionDeleteStyleTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>Styles: Normal (default), Default Paragraph Font, Heading1 (built in), Quote
    /// (custom, on Normal), MyQuote (custom, on Quote, next Quote, linked to MyQuoteChar), Plain
    /// (custom, on MyQuote, next MyQuote), Emph (custom character), Unused (custom).
    /// Body: Quote, MyQuote with an Emph run, MyQuote, unstyled. A numbering level tied to MyQuote.</summary>
    private static byte[] Document()
    {
        const string ns = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            void Write(OpenXmlPart part, string xml)
            {
                using var writer = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                writer.Write(xml);
            }
            Write(main, $"<w:document {ns}><w:body>"
                + "<w:p><w:pPr><w:pStyle w:val=\"Quote\"/></w:pPr><w:r><w:t>One</w:t></w:r></w:p>"
                + "<w:p><w:pPr><w:pStyle w:val=\"MyQuote\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"Emph\"/></w:rPr><w:t>Two</w:t></w:r></w:p>"
                + "<w:p><w:pPr><w:pStyle w:val=\"MyQuote\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"MyQuoteChar\"/></w:rPr><w:t>Three</w:t></w:r></w:p>"
                + "<w:p><w:r><w:t>Four</w:t></w:r></w:p>"
                + "</w:body></w:document>");
            Write(main.AddNewPart<StyleDefinitionsPart>(), $"<w:styles {ns}>"
                + "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/><w:qFormat/></w:style>"
                + "<w:style w:type=\"character\" w:default=\"1\" w:styleId=\"DefaultParagraphFont\"><w:name w:val=\"Default Paragraph Font\"/><w:semiHidden/></w:style>"
                + "<w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/><w:basedOn w:val=\"Normal\"/></w:style>"
                + "<w:style w:type=\"paragraph\" w:customStyle=\"1\" w:styleId=\"Quote\"><w:name w:val=\"Quote\"/><w:basedOn w:val=\"Normal\"/><w:rPr><w:i/></w:rPr></w:style>"
                + "<w:style w:type=\"paragraph\" w:customStyle=\"1\" w:styleId=\"MyQuote\"><w:name w:val=\"My Quote\"/><w:basedOn w:val=\"Quote\"/><w:next w:val=\"Quote\"/><w:link w:val=\"MyQuoteChar\"/><w:rPr><w:b/></w:rPr></w:style>"
                + "<w:style w:type=\"character\" w:customStyle=\"1\" w:styleId=\"MyQuoteChar\"><w:name w:val=\"My Quote Char\"/><w:basedOn w:val=\"DefaultParagraphFont\"/><w:link w:val=\"MyQuote\"/><w:rPr><w:b/></w:rPr></w:style>"
                + "<w:style w:type=\"paragraph\" w:customStyle=\"1\" w:styleId=\"Plain\"><w:name w:val=\"Plain\"/><w:basedOn w:val=\"MyQuote\"/><w:next w:val=\"MyQuote\"/></w:style>"
                + "<w:style w:type=\"character\" w:customStyle=\"1\" w:styleId=\"Emph\"><w:name w:val=\"Emph\"/><w:basedOn w:val=\"DefaultParagraphFont\"/><w:rPr><w:u w:val=\"single\"/></w:rPr></w:style>"
                + "<w:style w:type=\"paragraph\" w:customStyle=\"1\" w:styleId=\"Unused\"><w:name w:val=\"Unused\"/><w:basedOn w:val=\"Normal\"/></w:style>"
                + "</w:styles>");
            Write(main.AddNewPart<NumberingDefinitionsPart>(), $"<w:numbering {ns}>"
                + "<w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:pStyle w:val=\"MyQuote\"/><w:lvlText w:val=\"%1.\"/></w:lvl></w:abstractNum>"
                + "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>"
                + "</w:numbering>");
        }
        return stream.ToArray();
    }

    private static XElement Root(DocxSession session, System.Func<MainDocumentPart, OpenXmlPart> part)
    {
        using var ms = new MemoryStream(session.Save());
        using var doc = WordprocessingDocument.Open(ms, false);
        return part(doc.MainDocumentPart!).GetXDocument().Root!;
    }

    private static XElement? Style(DocxSession session, string id) =>
        Root(session, m => m.StyleDefinitionsPart!).Elements(W + "style").SingleOrDefault(s => (string?)s.Attribute(W + "styleId") == id);

    private static string?[] ParagraphStyles(DocxSession session) =>
        Root(session, m => m).Descendants(W + "p")
            .Select(p => (string?)p.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val")).ToArray();

    private static string?[] RunStyles(DocxSession session) =>
        Root(session, m => m).Descendants(W + "r")
            .Select(r => (string?)r.Element(W + "rPr")?.Element(W + "rStyle")?.Attribute(W + "val")).ToArray();

    [Fact]
    public void SD011_DeleteParagraphStyle_ParagraphsTakeItsBase_LinkedCharacterStyleGoes()
    {
        var session = new DocxSession(Document());
        var result = session.DeleteStyle("MyQuote");
        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal("MyQuote", result.StyleId);

        Assert.Null(Style(session, "MyQuote"));
        Assert.Null(Style(session, "MyQuoteChar"));
        Assert.Equal(new string?[] { "Quote", "Quote", "Quote", null }, ParagraphStyles(session));
        // The linked character style's run has no base to take but the default: no rStyle.
        Assert.Equal(new string?[] { null, "Emph", null, null }, RunStyles(session));
    }

    [Fact]
    public void SD012_StylesBasedOnIt_AreRebased_FollowingStyleDropped_NumberingTieDropped()
    {
        var session = new DocxSession(Document());
        Assert.True(session.DeleteStyle("MyQuote").Success);
        var plain = Style(session, "Plain")!;
        Assert.Equal("Quote", (string?)plain.Element(W + "basedOn")!.Attribute(W + "val"));
        Assert.Null(plain.Element(W + "next"));
        var lvl = Root(session, m => m.NumberingDefinitionsPart!).Descendants(W + "lvl").Single();
        Assert.Null(lvl.Element(W + "pStyle"));
        Assert.NotNull(lvl.Element(W + "lvlText"));
    }

    [Fact]
    public void SD013_DeleteCharacterStyle_RunsLoseIt_ADeletedBaseFallsToTheDefault()
    {
        var session = new DocxSession(Document());
        Assert.True(session.DeleteStyle("Emph").Success);
        Assert.Null(Style(session, "Emph"));
        Assert.Equal(new string?[] { null, null, "MyQuoteChar", null }, RunStyles(session));

        // Quote is based on the default: its paragraphs end up with no style reference at all.
        Assert.True(session.DeleteStyle("Quote").Success);
        Assert.Equal(new string?[] { null, "MyQuote", "MyQuote", null }, ParagraphStyles(session));
        Assert.Equal("Normal", (string?)Style(session, "MyQuote")!.Element(W + "basedOn")!.Attribute(W + "val"));
    }

    [Theory]
    [InlineData("Heading1")]
    [InlineData("Normal")]
    [InlineData("DefaultParagraphFont")]
    public void SD014_BuiltInAndDefaultStyles_AreRefused(string id)
    {
        var session = new DocxSession(Document());
        var result = session.DeleteStyle(id);
        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.StyleNotDeletable, result.Error!.Code);
        Assert.NotNull(Style(session, id));

        var unknown = session.DeleteStyle("Nope");
        Assert.Equal(EditErrorCode.UnknownStyle, unknown.Error!.Code);
        Assert.Contains("style_not_deletable", DocxSessionJson.Serialize(result));
    }

    [Fact]
    public void SD015_OneUndoStep_BringsEverythingBack_AndTheSavedDocumentIsValid()
    {
        var session = new DocxSession(Document());
        var before = ParagraphStyles(session);
        Assert.True(session.DeleteStyle("MyQuote").Success);
        using (var ms = new MemoryStream(session.Save()))
        using (var doc = WordprocessingDocument.Open(ms, false))
        {
            var errors = new OpenXmlValidator().Validate(doc).ToList();
            Assert.True(errors.Count == 0, string.Join("\n", errors.Select(e => e.Description)));
        }
        Assert.True(session.Undo());
        Assert.Equal(before, ParagraphStyles(session));
        Assert.NotNull(Style(session, "MyQuote"));
        Assert.NotNull(Style(session, "MyQuoteChar"));
        Assert.Equal("MyQuote", (string?)Style(session, "Plain")!.Element(W + "basedOn")!.Attribute(W + "val"));
        Assert.NotNull(Root(session, m => m.NumberingDefinitionsPart!).Descendants(W + "lvl").Single().Element(W + "pStyle"));
    }

    [Fact]
    public void SD016_InUse_MarksAppliedStyles_AndTheDefaultWhenAParagraphNamesNone()
    {
        var session = new DocxSession(Document());
        var inUse = session.ListStyles().Where(s => s.InUse).Select(s => s.Id).OrderBy(s => s).ToArray();
        Assert.Equal(new[] { "Emph", "MyQuote", "MyQuoteChar", "Normal", "Quote" }, inUse);

        var json = DocxSessionJson.SerializeStyles(session.ListStyles());
        Assert.Contains("\"id\":\"Quote\"", json);
        Assert.Contains("\"inUse\":true", json);

        Assert.True(session.DeleteStyle("Emph").Success);
        Assert.DoesNotContain(session.ListStyles(), s => s.Id == "Emph");
        Assert.False(session.ListStyles().Single(s => s.Id == "Unused").InUse);
    }
}
