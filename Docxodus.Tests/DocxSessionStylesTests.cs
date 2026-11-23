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
/// Creating and changing paragraph and character styles with <see cref="DocxSession.SetStyle"/>
/// (Word's Create New Style and Modify Style). Test IDs use the SD0xx range.
/// </summary>
public class DocxSessionStylesTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>Normal (default), Heading1 (bold, keep with next, 16 pt) linked to Heading1Char,
    /// Quote (custom, based on Normal), the default character style and a latent "heading 2".
    /// The first paragraph is a Heading 1, the second Normal.</summary>
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
                + "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>"
                + "<w:p><w:r><w:t>Body</w:t></w:r></w:p>"
                + "</w:body></w:document>");
            Write(main.AddNewPart<StyleDefinitionsPart>(), $"<w:styles {ns}>"
                + "<w:latentStyles><w:lsdException w:name=\"heading 2\" w:uiPriority=\"9\" w:qFormat=\"1\"/></w:latentStyles>"
                + "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/><w:qFormat/></w:style>"
                + "<w:style w:type=\"character\" w:default=\"1\" w:styleId=\"DefaultParagraphFont\"><w:name w:val=\"Default Paragraph Font\"/><w:uiPriority w:val=\"1\"/><w:semiHidden/></w:style>"
                + "<w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/><w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/><w:link w:val=\"Heading1Char\"/><w:qFormat/>"
                + "<w:pPr><w:keepNext/></w:pPr><w:rPr><w:b/><w:sz w:val=\"32\"/></w:rPr></w:style>"
                + "<w:style w:type=\"character\" w:customStyle=\"1\" w:styleId=\"Heading1Char\"><w:name w:val=\"Heading 1 Char\"/><w:basedOn w:val=\"DefaultParagraphFont\"/><w:link w:val=\"Heading1\"/>"
                + "<w:rPr><w:b/><w:sz w:val=\"32\"/></w:rPr></w:style>"
                + "<w:style w:type=\"paragraph\" w:customStyle=\"1\" w:styleId=\"Quote\"><w:name w:val=\"Quote\"/><w:basedOn w:val=\"Normal\"/></w:style>"
                + "</w:styles>");
        }
        return stream.ToArray();
    }

    private static string[] Paragraphs(DocxSession session) =>
        session.Project().AnchorIndex.Values.Where(t => t.Anchor.Kind is "p" or "h").Select(t => t.Anchor.Id).ToArray();

    private static XElement Styles(DocxSession session)
    {
        using var ms = new MemoryStream(session.Save());
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.StyleDefinitionsPart!.GetXDocument().Root!;
    }

    private static XElement Style(DocxSession session, string id) =>
        Styles(session).Elements(W + "style").Single(s => (string?)s.Attribute(W + "styleId") == id);

    private static RunFormattingInfo EffectiveRun(DocxSession session, int paragraph) =>
        session.GetFormatting(Paragraphs(session)[paragraph])!.Runs[0].Effective;

    [Fact]
    public void SD001_NewParagraphStyle_TakesWordsDefaults_AndIsUsable()
    {
        var session = new DocxSession(Document());
        var result = session.SetStyle(new StyleDefinitionOp
        {
            Name = "Block Text 2",
            Paragraph = new ParagraphFormatOp { Alignment = ParagraphAlignment.Center, SpacingAfter = 240 },
            Run = new FormatOp { Italic = true, FontSizePts = 13 },
        });
        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal("BlockText2", result.StyleId);

        var style = Style(session, "BlockText2");
        Assert.Equal("paragraph", (string?)style.Attribute(W + "type"));
        Assert.Equal("1", (string?)style.Attribute(W + "customStyle"));
        Assert.Equal(new[] { "name", "basedOn", "next", "qFormat", "pPr", "rPr" },
            style.Elements().Select(e => e.Name.LocalName));
        Assert.Equal("Normal", (string?)style.Element(W + "basedOn")!.Attribute(W + "val"));
        Assert.Equal("BlockText2", (string?)style.Element(W + "next")!.Attribute(W + "val"));

        var info = session.ListStyles().Single(s => s.Id == "BlockText2");
        Assert.Equal("Block Text 2", info.Name);
        Assert.True(info.QuickFormat);

        Assert.True(session.SetParagraphStyle(Paragraphs(session)[1], "BlockText2").Success);
        var formatting = session.GetFormatting(Paragraphs(session)[1])!;
        Assert.Equal(ParagraphAlignment.Center, formatting.EffectiveParagraph.Alignment);
        Assert.Equal(240, formatting.EffectiveParagraph.SpacingAfterTwips);
        Assert.True(formatting.Runs[0].Effective.Italic);
        Assert.Equal(13, formatting.Runs[0].Effective.FontSizePts);
    }

    [Fact]
    public void SD002_ModifyBuiltInStyle_ChangesItsParagraphs_AndItsLinkedCharacterStyle()
    {
        var session = new DocxSession(Document());
        var result = session.SetStyle(new StyleDefinitionOp
        {
            StyleId = "Heading1",
            Run = new FormatOp { FontSizePts = 20, Color = "1F4E79" },
        });
        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal("Heading1", result.StyleId);

        var run = EffectiveRun(session, 0);
        Assert.Equal(20, run.FontSizePts);
        Assert.True(run.Bold);
        var linked = Style(session, "Heading1Char").Element(W + "rPr")!;
        Assert.Equal("40", (string?)linked.Element(W + "sz")!.Attribute(W + "val"));
        Assert.Equal("1F4E79", (string?)linked.Element(W + "color")!.Attribute(W + "val"));
        // The Normal paragraph is untouched.
        Assert.NotEqual(20, EffectiveRun(session, 1).FontSizePts);
    }

    [Fact]
    public void SD003_TurningOffWhatTheBasedOnStyleTurnsOn_WritesAnExplicitOff()
    {
        var session = new DocxSession(Document());
        var created = session.SetStyle(new StyleDefinitionOp
        {
            Name = "Plain Heading",
            BasedOn = "Heading1",
            Run = new FormatOp { Bold = false },
            Paragraph = new ParagraphFormatOp { KeepNext = false },
        });
        Assert.True(created.Success, created.Error?.Message);
        var style = Style(session, created.StyleId!);
        Assert.Equal("0", (string?)style.Element(W + "rPr")!.Element(W + "b")!.Attribute(W + "val"));
        Assert.Equal("0", (string?)style.Element(W + "pPr")!.Element(W + "keepNext")!.Attribute(W + "val"));

        Assert.True(session.SetParagraphStyle(Paragraphs(session)[1], created.StyleId!).Success);
        var formatting = session.GetFormatting(Paragraphs(session)[1])!;
        Assert.NotEqual(true, formatting.Runs[0].Effective.Bold);
        Assert.NotEqual(true, formatting.EffectiveParagraph.KeepNext);

        // Off where nothing turns it on writes nothing.
        var quote = session.SetStyle(new StyleDefinitionOp { StyleId = "Quote", Run = new FormatOp { Bold = false } });
        Assert.True(quote.Success, quote.Error?.Message);
        Assert.Null(Style(session, "Quote").Element(W + "rPr"));
    }

    [Fact]
    public void SD004_NewCharacterStyle_IsBasedOnTheDefaultCharacterStyle_WithNoFollowingStyle()
    {
        var session = new DocxSession(Document());
        var result = session.SetStyle(new StyleDefinitionOp
        {
            Name = "Key Term",
            Type = "character",
            Run = new FormatOp { Bold = true, SmallCaps = true },
        });
        Assert.True(result.Success, result.Error?.Message);
        var style = Style(session, "KeyTerm");
        Assert.Equal("character", (string?)style.Attribute(W + "type"));
        Assert.Equal("DefaultParagraphFont", (string?)style.Element(W + "basedOn")!.Attribute(W + "val"));
        Assert.Null(style.Element(W + "next"));
        Assert.NotNull(style.Element(W + "rPr")!.Element(W + "smallCaps"));
    }

    [Fact]
    public void SD005_RenameBasedOnNextAndGallery_OnACustomStyle()
    {
        var session = new DocxSession(Document());
        var result = session.SetStyle(new StyleDefinitionOp
        {
            StyleId = "Quote",
            Name = "Pull Quote",
            BasedOn = "Heading1",
            Next = "Normal",
            QuickFormat = true,
        });
        Assert.True(result.Success, result.Error?.Message);
        var style = Style(session, "Quote");
        Assert.Equal("Pull Quote", (string?)style.Element(W + "name")!.Attribute(W + "val"));
        Assert.Equal("Heading1", (string?)style.Element(W + "basedOn")!.Attribute(W + "val"));
        Assert.Equal("Normal", (string?)style.Element(W + "next")!.Attribute(W + "val"));
        Assert.NotNull(style.Element(W + "qFormat"));

        Assert.True(session.SetStyle(new StyleDefinitionOp { StyleId = "Quote", BasedOn = "", QuickFormat = false }).Success);
        style = Style(session, "Quote");
        Assert.Null(style.Element(W + "basedOn"));
        Assert.Null(style.Element(W + "qFormat"));
    }

    [Theory]
    [InlineData("{\"name\":\"\"}")]
    [InlineData("{\"name\":\"normal\"}")]
    [InlineData("{\"name\":\"Heading 2\"}")]
    [InlineData("{\"styleId\":\"Heading1\",\"name\":\"Big Title\"}")]
    [InlineData("{\"styleId\":\"Normal\",\"basedOn\":\"Quote\"}")]
    [InlineData("{\"styleId\":\"Normal\",\"basedOn\":\"Normal\"}")]
    [InlineData("{\"styleId\":\"Quote\",\"basedOn\":\"Heading1Char\"}")]
    [InlineData("{\"styleId\":\"Quote\",\"type\":\"character\"}")]
    [InlineData("{\"name\":\"Mark\",\"type\":\"character\",\"paragraph\":{\"alignment\":\"center\"}}")]
    [InlineData("{\"name\":\"Mark\",\"type\":\"character\",\"next\":\"Normal\"}")]
    [InlineData("{\"name\":\"Table Look\",\"type\":\"table\"}")]
    [InlineData("{\"styleId\":\"Quote\",\"run\":{\"runStyle\":\"Heading1Char\"}}")]
    [InlineData("{\"styleId\":\"Quote\",\"run\":{\"vertAlign\":\"sideways\"}}")]
    [InlineData("{\"styleId\":\"Quote\",\"run\":{\"highlight\":\"yellow\"}}")]
    [InlineData("{\"styleId\":\"Quote\",\"paragraph\":{\"firstLineIndent\":10,\"hangingIndent\":10}}")]
    public void SD006_Refusals_LeaveTheStylesAlone_AndRecordNoUndo(string opJson)
    {
        var session = new DocxSession(Document());
        var before = Styles(session).ToString();
        var result = session.SetStyle(DocxSessionJson.ParseStyleDefinitionOp(opJson));
        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.InvalidStyleDefinition, result.Error!.Code);
        Assert.Equal(before, Styles(session).ToString());
        Assert.False(session.Undo());
    }

    [Fact]
    public void SD007_UnknownStyles_AreUnknownStyle()
    {
        var session = new DocxSession(Document());
        Assert.Equal(EditErrorCode.UnknownStyle,
            session.SetStyle(new StyleDefinitionOp { StyleId = "Missing", Run = new FormatOp { Bold = true } }).Error!.Code);
        Assert.Equal(EditErrorCode.UnknownStyle,
            session.SetStyle(new StyleDefinitionOp { StyleId = "Quote", BasedOn = "Missing" }).Error!.Code);
        Assert.Equal(EditErrorCode.UnknownStyle,
            session.SetStyle(new StyleDefinitionOp { StyleId = "Quote", Next = "Missing" }).Error!.Code);
    }

    [Fact]
    public void SD008_OneUndoStep_RemovesANewStyle_AndRestoresAChangedOne()
    {
        var session = new DocxSession(Document());
        Assert.True(session.SetStyle(new StyleDefinitionOp { Name = "Quote" + "d", Run = new FormatOp { Bold = true } }).Success);
        Assert.True(session.Undo());
        Assert.DoesNotContain(session.ListStyles(), s => s.Id == "Quoted");

        Assert.True(session.SetStyle(new StyleDefinitionOp { StyleId = "Heading1", Run = new FormatOp { FontSizePts = 30 } }).Success);
        Assert.Equal(30, EffectiveRun(session, 0).FontSizePts);
        Assert.True(session.Undo());
        Assert.Equal(16, EffectiveRun(session, 0).FontSizePts);
    }

    [Fact]
    public void SD009_IdsAreMadeFromAsciiLettersAndDigits_AndNumberedWhenTaken()
    {
        var session = new DocxSession(Document());
        Assert.Equal("Ttulo3", session.SetStyle(new StyleDefinitionOp { Name = "Título 3" }).StyleId);
        // Another name whose id comes out the same gets the next number.
        Assert.Equal("Ttulo31", session.SetStyle(new StyleDefinitionOp { Name = "Tétulo 3" }).StyleId);
        Assert.Equal("Style", session.SetStyle(new StyleDefinitionOp { Name = "ñ" }).StyleId);
    }

    [Fact]
    public void SD010_TheSavedDocumentIsValid_AndTheResultCarriesTheStyleId()
    {
        var session = new DocxSession(Document());
        Assert.True(session.SetStyle(new StyleDefinitionOp
        {
            Name = "Callout",
            Paragraph = new ParagraphFormatOp { LeftIndent = 720, KeepLines = true, LineSpacing = 360, LineSpacingRule = LineSpacingRule.Auto },
            Run = new FormatOp { Underline = true, FontFamily = "Georgia", Color = "C00000" },
        }).Success);
        using var ms = new MemoryStream(session.Save());
        using var doc = WordprocessingDocument.Open(ms, false);
        var errors = new OpenXmlValidator().Validate(doc.MainDocumentPart!.StyleDefinitionsPart!).ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors.Select(e => e.Description)));

        var json = DocxSessionJson.Serialize(
            session.SetStyle(new StyleDefinitionOp { StyleId = "Callout", Run = new FormatOp { Italic = true } }));
        Assert.Contains("\"styleId\":\"Callout\"", json);
    }
}
