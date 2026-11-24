using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary><see cref="DocxSession.ExpandRangeToFields"/> (a range reaching into a field grows to
/// the whole field, as Word's selection does), <see cref="DocxSession.RepairFields"/> (markers left
/// without partners taken out) and an editor render that survives such markers.</summary>
public class DocxSessionFieldRangesTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static string Run(string inner) => $"<w:r>{inner}</w:r>";
    private static string Char(string type) => Run($"<w:fldChar w:fldCharType=\"{type}\"/>");
    private static string Instr(string code) => Run($"<w:instrText xml:space=\"preserve\"> {code} </w:instrText>");
    private static string Text(string text) => Run($"<w:t xml:space=\"preserve\">{text}</w:t>");
    private static string PageRef(string bookmark, string page) =>
        Char("begin") + Instr($"PAGEREF {bookmark} \\h") + Char("separate") + Text(page) + Char("end");

    /// <summary>"Before the table", a contents table of two entries (the field begins at the
    /// first entry's start and ends at the second's end, as Word writes it), "After it".</summary>
    private const string Contents =
        "<w:p>" + "<w:r><w:t>Before the table</w:t></w:r>" + "</w:p>"
        + "<w:p>" + "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> TOC \\o \"1-3\" \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
        + "<w:r><w:t>Alpha</w:t></w:r><w:r><w:tab/></w:r>" + "PAGEREF_A" + "</w:p>"
        + "<w:p>" + "<w:r><w:t>Beta</w:t></w:r><w:r><w:tab/></w:r>" + "PAGEREF_B" + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" + "</w:p>"
        + "<w:p>" + "<w:r><w:t>After it</w:t></w:r>" + "</w:p>";

    private static byte[] Build(string bodyXml)
    {
        var xml = $"<w:document xmlns:w=\"{W.NamespaceName}\"><w:body>{bodyXml}"
            + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>";
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.PutXDocument(XDocument.Parse(xml));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(new XElement(W + "styles")));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(new XDocument(new XElement(W + "settings")));
        }
        return stream.ToArray();
    }

    private static DocxSession WithContents() =>
        new(Build(Contents.Replace("PAGEREF_A", PageRef("_Toc1", "1")).Replace("PAGEREF_B", PageRef("_Toc2", "2"))));

    private static string[] Paragraphs(DocxSession session) =>
        session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Kind is "p" or "h" && t.Anchor.Scope == "body")
            .Select(t => t.Anchor.Id).ToArray();

    private static XElement Body(DocxSession session)
    {
        using var stream = new MemoryStream(session.Save());
        using var doc = WordprocessingDocument.Open(stream, false);
        return XDocument.Parse(new StreamReader(doc.MainDocumentPart!.GetStream()).ReadToEnd()).Root!.Element(W + "body")!;
    }

    private static string[] Markers(DocxSession session) =>
        Body(session).Descendants(W + "fldChar").Select(f => (string)f.Attribute(W + "fldCharType")!).ToArray();

    [Fact]
    public void FR001_ARangeFromBeforeTheContentsIntoThem_GrowsToTheFieldsEnd()
    {
        using var session = WithContents();
        var p = Paragraphs(session);
        var grown = session.ExpandRangeToFields(new DocumentRange(p[0], 3, p[1], 2));
        Assert.NotNull(grown);
        Assert.Equal((p[0], 3, p[2]), (grown!.StartAnchorId, grown.StartOffset, grown.EndAnchorId));
        Assert.Equal("Beta\t2".Length, grown.EndOffset);
    }

    [Fact]
    public void FR002_ARangeFromInsideTheContentsPastThem_GrowsToTheFieldsBegin()
    {
        using var session = WithContents();
        var p = Paragraphs(session);
        var grown = session.ExpandRangeToFields(new DocumentRange(p[1], 2, p[3], 3));
        Assert.Equal(new DocumentRange(p[1], 0, p[3], 3), grown);
    }

    [Fact]
    public void FR003_ARangeInsideTheContents_OrCoveringAWholeField_StaysAsItIs()
    {
        using var session = WithContents();
        var p = Paragraphs(session);
        // (entries' text, the first entry's page reference whole)
        var inside = new DocumentRange(p[1], 2, p[2], 2);
        Assert.Equal(inside, session.ExpandRangeToFields(inside));
        var whole = new DocumentRange(p[0], 0, p[3], 3);
        Assert.Equal(whole, session.ExpandRangeToFields(whole));
        // (ending where the page reference begins does not reach into it)
        var before = new DocumentRange(p[1], 1, p[1], "Alpha\t".Length);
        Assert.Equal(before, session.ExpandRangeToFields(before));
    }

    [Fact]
    public void FR004_RepairFields_TakesOutAnEndWithoutItsBegin_InOneUndoStep()
    {
        // (the contents table's begin is gone, as a delete across it left it before)
        var body = "<w:p><w:r><w:t>Before</w:t></w:r></w:p>"
            + "<w:p><w:r><w:t>Beta</w:t></w:r>" + PageRef("_Toc2", "2") + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>";
        using var session = new DocxSession(Build(body));
        var p = Paragraphs(session);
        int undo = session.UndoCount;
        var result = session.RepairFields(p[1]);
        Assert.True(result.Success);
        Assert.Equal(new[] { "begin", "separate", "end" }, Markers(session));
        Assert.Contains(result.Modified, a => a.Id == p[1]);
        Assert.Equal(undo + 1, session.UndoCount);
        // nothing left to repair: no undo step
        Assert.True(session.RepairFields(p[1]).Success);
        Assert.Equal(undo + 1, session.UndoCount);
        Assert.True(session.Undo());
        Assert.Equal(new[] { "begin", "separate", "end", "end" }, Markers(session));
    }

    [Fact]
    public void FR005_RepairFields_TakesOutABeginNeverEnded_WithItsCode_KeepingItsResultText()
    {
        var body = "<w:p>" + Char("begin") + Instr("TOC \\o") + Char("separate") + Text("Alpha") + "</w:p>"
            + "<w:p>" + Text("Next") + "</w:p>";
        using var session = new DocxSession(Build(body));
        var p = Paragraphs(session);
        Assert.True(session.RepairFields(p[1]).Success);
        var b = Body(session);
        Assert.Empty(b.Descendants(W + "fldChar"));
        Assert.Empty(b.Descendants(W + "instrText"));
        Assert.Equal("AlphaNext", string.Concat(b.Descendants(W + "t").Select(t => t.Value)));
    }

    [Fact]
    public void FR006_EditorRender_PassesOverAnEndWithoutItsBegin()
    {
        var body = "<w:p><w:r><w:t>Before</w:t></w:r></w:p>"
            + "<w:p><w:r><w:t>Beta</w:t></w:r>" + PageRef("_Toc2", "2") + Char("end") + Char("separate") + "</w:p>"
            + "<w:p><w:r><w:t>After</w:t></w:r></w:p>";
        int handle = SessionRegistry.OpenSession(Build(body), new DocxSessionSettings());
        try
        {
            var html = DocxSessionOps.RenderEditorHtml(handle, "{\"paginated\":true,\"referenceFields\":true}");
            Assert.Contains("Beta", html);
            Assert.Contains("After", html);
        }
        finally { SessionRegistry.CloseSession(handle); }
    }
}
