using System;
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
/// Copying styles from another document or template with <see cref="DocxSession.CopyStyles"/>
/// (Word's Organizer, from a file into the document). Test IDs use the CS0xx range.
/// </summary>
public class DocxSessionCopyStylesTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const string Ns = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";

    private static byte[] Package(string body, string styles, string? numbering = null,
        DocumentFormat.OpenXml.WordprocessingDocumentType type = DocumentFormat.OpenXml.WordprocessingDocumentType.Document)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, type))
        {
            var main = doc.AddMainDocumentPart();
            void Write(OpenXmlPart part, string xml)
            {
                using var writer = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                writer.Write(xml);
            }
            Write(main, $"<w:document {Ns}><w:body>{body}</w:body></w:document>");
            Write(main.AddNewPart<StyleDefinitionsPart>(), $"<w:styles {Ns}>{styles}</w:styles>");
            if (numbering is not null) Write(main.AddNewPart<NumberingDefinitionsPart>(), $"<w:numbering {Ns}>{numbering}</w:numbering>");
        }
        return stream.ToArray();
    }

    private const string NormalStyle =
        "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/><w:qFormat/></w:style>";

    /// <summary>The document: Normal (default), heading 1 (bold), Quote (custom). Paragraphs: a
    /// Heading 1, then Normal text. No numbering part.</summary>
    private static byte[] Target() => Package(
        "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>"
        + "<w:p><w:r><w:t>Body</w:t></w:r></w:p>",
        NormalStyle
        + "<w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/><w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/><w:qFormat/>"
        + "<w:rPr><w:b/><w:sz w:val=\"32\"/></w:rPr></w:style>"
        + "<w:style w:type=\"paragraph\" w:customStyle=\"1\" w:styleId=\"Quote\"><w:name w:val=\"Quote\"/><w:basedOn w:val=\"Normal\"/></w:style>");

    /// <summary>The other document: its own Normal (red, default) and heading 1 (italic, 20 pt, id
    /// H1); Fancy based on Base (which the document lacks), followed by Missing (which nobody
    /// has), linked to Fancy Char; a "Block Quote" whose id is the document's Quote's; a
    /// numbered style; a default paragraph style named Body Text.</summary>
    private static byte[] Source(DocumentFormat.OpenXml.WordprocessingDocumentType type = DocumentFormat.OpenXml.WordprocessingDocumentType.Document) => Package(
        "<w:p><w:r><w:t>x</w:t></w:r></w:p>",
        "<w:style w:type=\"paragraph\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/><w:rPr><w:color w:val=\"FF0000\"/></w:rPr></w:style>"
        + "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"BodyText\"><w:name w:val=\"Body Text\"/><w:basedOn w:val=\"Normal\"/></w:style>"
        + "<w:style w:type=\"paragraph\" w:styleId=\"H1\"><w:name w:val=\"Heading 1\"/><w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/>"
        + "<w:rPr><w:i/><w:sz w:val=\"40\"/></w:rPr></w:style>"
        + "<w:style w:type=\"paragraph\" w:customStyle=\"1\" w:styleId=\"Base\"><w:name w:val=\"Base\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:jc w:val=\"center\"/></w:pPr></w:style>"
        + "<w:style w:type=\"paragraph\" w:customStyle=\"1\" w:styleId=\"Fancy\"><w:name w:val=\"Fancy\"/><w:basedOn w:val=\"Base\"/><w:next w:val=\"Missing\"/><w:link w:val=\"FancyChar\"/>"
        + "<w:rPr><w:caps/></w:rPr></w:style>"
        + "<w:style w:type=\"character\" w:customStyle=\"1\" w:styleId=\"FancyChar\"><w:name w:val=\"Fancy Char\"/><w:link w:val=\"Fancy\"/><w:rPr><w:caps/></w:rPr></w:style>"
        + "<w:style w:type=\"paragraph\" w:customStyle=\"1\" w:styleId=\"Quote\"><w:name w:val=\"Block Quote\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:ind w:left=\"720\"/></w:pPr></w:style>"
        + "<w:style w:type=\"paragraph\" w:customStyle=\"1\" w:styleId=\"Steps\"><w:name w:val=\"Steps\"/><w:basedOn w:val=\"Normal\"/>"
        + "<w:pPr><w:numPr><w:numId w:val=\"7\"/></w:numPr></w:pPr></w:style>",
        "<w:abstractNum w:abstractNumId=\"3\"><w:nsid w:val=\"1A2B3C4D\"/><w:multiLevelType w:val=\"singleLevel\"/>"
        + "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"upperRoman\"/><w:pStyle w:val=\"Steps\"/><w:lvlText w:val=\"%1.\"/><w:lvlJc w:val=\"left\"/></w:lvl></w:abstractNum>"
        + "<w:num w:numId=\"7\"><w:abstractNumId w:val=\"3\"/></w:num>",
        type);

    private static XElement Root(DocxSession session, Func<MainDocumentPart, OpenXmlPart?> part)
    {
        using var ms = new MemoryStream(session.Save());
        using var doc = WordprocessingDocument.Open(ms, false);
        return part(doc.MainDocumentPart!)!.GetXDocument().Root!;
    }

    private static XElement Styles(DocxSession session) => Root(session, m => m.StyleDefinitionsPart);

    private static XElement? Style(XElement styles, string id) =>
        styles.Elements(W + "style").SingleOrDefault(s => (string?)s.Attribute(W + "styleId") == id);

    private static string? Val(XElement? style, string child) =>
        (string?)style?.Element(W + child)?.Attribute(W + "val");

    private static string[] Paragraphs(DocxSession session) =>
        session.Project().AnchorIndex.Values.Where(t => t.Anchor.Kind is "p" or "h").Select(t => t.Anchor.Id).ToArray();

    [Fact]
    public void CS001_ListStylesIn_ListsTheOtherDocumentsStyles_AndRefusesNonPackages()
    {
        var styles = DocxSession.ListStylesIn(Source());
        Assert.Contains(styles, s => s.Id == "H1" && s.Name == "Heading 1" && s.Type == "paragraph");
        Assert.Contains(styles, s => s.Id == "FancyChar" && s.Type == "character");
        Assert.Equal(8, styles.Count);

        Assert.Throws<ArgumentException>(() => DocxSession.ListStylesIn(new byte[] { 1, 2, 3 }));
        Assert.Throws<ArgumentException>(() => DocxSession.ListStylesIn(Array.Empty<byte>()));
        Assert.Contains("InvalidStyleSource", DocxSessionOps.ListStylesIn("AQID"));
        Assert.StartsWith("[", DocxSessionOps.ListStylesIn(Convert.ToBase64String(Source())));
    }

    [Fact]
    public void CS002_SameName_ReplacesTheStyle_KeepingItsIdAndItsParagraphs()
    {
        var session = new DocxSession(Target());
        var result = session.CopyStyles(Source(), new[] { "H1" });
        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal("Heading1", result.StyleId);

        var styles = Styles(session);
        Assert.Null(Style(styles, "H1"));
        var heading = Style(styles, "Heading1")!;
        Assert.Equal("Heading 1", Val(heading, "name"));
        Assert.NotNull(heading.Element(W + "rPr")!.Element(W + "i"));
        Assert.Null(heading.Element(W + "rPr")!.Element(W + "b"));
        // Its based-on Normal is this document's: not replaced by the other's red one.
        Assert.Equal("Normal", Val(heading, "basedOn"));
        Assert.Null(Style(styles, "Normal")!.Element(W + "rPr"));

        var run = session.GetFormatting(Paragraphs(session)[0])!.Runs[0].Effective;
        Assert.True(run.Italic);
        Assert.Equal(20, run.FontSizePts);
    }

    [Fact]
    public void CS003_MissingBasedOnAndLinkedStylesComeAlong_AMissingNextIsDropped()
    {
        var session = new DocxSession(Target());
        Assert.True(session.CopyStyles(Source(), new[] { "Fancy" }).Success);
        var styles = Styles(session);
        var fancy = Style(styles, "Fancy")!;
        Assert.Equal("Base", Val(fancy, "basedOn"));
        Assert.Equal("FancyChar", Val(fancy, "link"));
        Assert.Null(fancy.Element(W + "next"));
        Assert.NotNull(Style(styles, "Base"));
        Assert.Equal("Fancy", Val(Style(styles, "FancyChar"), "link"));
        Assert.Equal("Normal", Val(Style(styles, "Base"), "basedOn"));
        Assert.Null(Style(styles, "Normal")!.Element(W + "rPr"));
    }

    [Fact]
    public void CS004_AnIdTakenByAnotherStyle_GetsANewId_AndTheOtherStyleStays()
    {
        var session = new DocxSession(Target());
        var result = session.CopyStyles(Source(), new[] { "Quote" });
        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal("BlockQuote", result.StyleId);
        var styles = Styles(session);
        Assert.Equal("Block Quote", Val(Style(styles, "BlockQuote"), "name"));
        Assert.Equal("Quote", Val(Style(styles, "Quote"), "name"));
    }

    [Fact]
    public void CS005_TheDocumentsDefaultStaysTheDefault()
    {
        var session = new DocxSession(Target());
        Assert.True(session.CopyStyles(Source(), new[] { "BodyText", "Normal" }).Success);
        var styles = Styles(session);
        Assert.Null(Style(styles, "BodyText")!.Attribute(W + "default"));
        var normal = Style(styles, "Normal")!;
        Assert.Equal("1", (string?)normal.Attribute(W + "default"));
        Assert.Equal("FF0000", (string?)normal.Element(W + "rPr")!.Element(W + "color")!.Attribute(W + "val"));
        Assert.Single(styles.Elements(W + "style").Where(s => (string?)s.Attribute(W + "type") == "paragraph"
            && (string?)s.Attribute(W + "default") == "1"));
    }

    [Fact]
    public void CS006_AStylesListIsCopiedAsAListOfItsOwn()
    {
        // The document has a list already, with the source's nsid.
        var target = Package("<w:p><w:r><w:t>x</w:t></w:r></w:p>", NormalStyle,
            "<w:abstractNum w:abstractNumId=\"0\"><w:nsid w:val=\"1A2B3C4D\"/><w:lvl w:ilvl=\"0\"><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"*\"/></w:lvl></w:abstractNum>"
            + "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>");
        var session = new DocxSession(target);
        Assert.True(session.CopyStyles(Source(), new[] { "Steps" }).Success);

        var steps = Style(Styles(session), "Steps")!;
        var numId = (string?)steps.Element(W + "pPr")!.Element(W + "numPr")!.Element(W + "numId")!.Attribute(W + "val");
        Assert.Equal("2", numId);
        var numbering = Root(session, m => m.NumberingDefinitionsPart);
        var num = numbering.Elements(W + "num").Single(n => (string?)n.Attribute(W + "numId") == "2");
        var abstractId = (string?)num.Element(W + "abstractNumId")!.Attribute(W + "val");
        Assert.Equal("1", abstractId);
        var abstractNum = numbering.Elements(W + "abstractNum").Single(a => (string?)a.Attribute(W + "abstractNumId") == "1");
        Assert.Equal("upperRoman", (string?)abstractNum.Descendants(W + "numFmt").First().Attribute(W + "val"));
        Assert.NotEqual("1A2B3C4D", (string?)abstractNum.Element(W + "nsid")!.Attribute(W + "val"));
        Assert.Equal("Steps", (string?)abstractNum.Descendants(W + "pStyle").Single().Attribute(W + "val"));
        Assert.Equal(2, numbering.Elements(W + "abstractNum").Count());
    }

    [Fact]
    public void CS007_OneUndo_PutsEverythingBack_EvenANewNumberingPart()
    {
        var session = new DocxSession(Target());
        var before = Styles(session).ToString();
        Assert.True(session.CopyStyles(Source(), new[] { "H1", "Steps", "Fancy" }).Success);
        Assert.NotNull(Root(session, m => m.NumberingDefinitionsPart));
        Assert.True(session.Undo());
        Assert.Equal(before, Styles(session).ToString());
        using var ms = new MemoryStream(session.Save());
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.Null(doc.MainDocumentPart!.NumberingDefinitionsPart);
        Assert.True(session.Redo());
        Assert.NotNull(Style(Styles(session), "Steps"));
    }

    [Fact]
    public void CS008_Refusals()
    {
        var session = new DocxSession(Target());
        Assert.Equal(EditErrorCode.UnknownStyle, session.CopyStyles(Source(), new[] { "Nope" }).Error!.Code);
        Assert.Equal(EditErrorCode.InvalidStyleDefinition, session.CopyStyles(Source(), Array.Empty<string>()).Error!.Code);
        Assert.Equal(EditErrorCode.InvalidStyleSource, session.CopyStyles(new byte[] { 1, 2 }, new[] { "H1" }).Error!.Code);
        Assert.Equal(0, session.Version);
    }

    [Fact]
    public void CS009_FromATemplate_AndTheResultIsValid()
    {
        var session = new DocxSession(Target());
        var source = Source(DocumentFormat.OpenXml.WordprocessingDocumentType.Template);
        Assert.Contains(DocxSession.ListStylesIn(source), s => s.Id == "Fancy");
        var all = DocxSession.ListStylesIn(source).Select(s => s.Id).ToArray();
        Assert.True(session.CopyStyles(source, all).Success);

        using var ms = new MemoryStream(session.Save());
        using var doc = WordprocessingDocument.Open(ms, false);
        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc.MainDocumentPart!.StyleDefinitionsPart!)
            .Concat(validator.Validate(doc.MainDocumentPart!.NumberingDefinitionsPart!)).ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors.Select(e => e.Description)));

        var json = DocxSessionJson.Serialize(session.CopyStyles(source, new[] { "Base" }));
        Assert.Contains("\"success\":true", json);
    }
}
