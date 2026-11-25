using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;
using static Docxodus.Tests.DocxSessionUpdateFieldsTests;

namespace Docxodus.Tests;

/// <summary>Word's Insert Caption (<see cref="DocxSession.InsertCaption"/>), the captions a document
/// holds (<see cref="DocxSession.ListCaptions"/>), cross-reference bookmarks over part of a caption,
/// and a table of figures that <see cref="DocxSession.UpdateFields"/> fills.</summary>
public class DocxSessionCaptionsTests
{
    private static readonly XNamespace W = DocxSessionUpdateFieldsTests.W;

    private static string TableXml(string cellText) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>"
        + $"<w:tr><w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr>{P(R(cellText))}</w:tc></w:tr></w:tbl>";

    private static string Anchor(DocxSession session, string kind, int index = 0) =>
        session.Project().AnchorIndex.Values.Where(t => t.Anchor.Kind == kind && t.Anchor.Scope == "body")
            .Select(t => t.Anchor.Id).ElementAt(index);

    private static XElement Styles(DocxSession session)
    {
        using var stream = new MemoryStream(session.Save());
        using var doc = WordprocessingDocument.Open(stream, false);
        return XDocument.Parse(new StreamReader(doc.MainDocumentPart!.StyleDefinitionsPart!.GetStream()).ReadToEnd()).Root!;
    }

    private static string? StyleOf(XElement paragraph) =>
        (string?)paragraph.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val");

    private static string Caption(string label, string number, string text = "") =>
        P(R(label + " ") + Field($" SEQ {label} \\* ARABIC ", number) + (text.Length > 0 ? R(text) : ""), "Caption");

    [Fact]
    public void CP001_BelowAParagraph_InWordsCaptionStyle_OneUndo()
    {
        using var session = Open(P(R("A picture")) + P(R("After")));
        var before = Body(session).ToString();

        var result = session.InsertCaption(BodyParagraphs(session)[0], new CaptionOptions { Text = ": A cat" });

        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal(new[] { "A picture", "Figure 1: A cat", "After" }, Lines(session));
        var caption = Body(session).Elements(W + "p").ElementAt(1);
        Assert.Equal("Caption", StyleOf(caption));
        Assert.Contains(" SEQ Figure \\* ARABIC ", caption.Descendants(W + "instrText").Single().Value);
        Assert.Null(caption.Element(W + "pPr")!.Element(W + "keepNext"));
        var style = Styles(session).Elements(W + "style").Single(s => (string?)s.Attribute(W + "styleId") == "Caption");
        Assert.Equal("caption", (string?)style.Element(W + "name")!.Attribute(W + "val"));
        Assert.Single(result.Created);

        Assert.True(session.Undo());
        Assert.Equal(before, Body(session).ToString());
    }

    [Fact]
    public void CP002_ATable_TakesItsCaptionAbove_KeptWithIt()
    {
        using var session = Open(P(R("Intro")) + TableXml("Cell"));

        Assert.True(session.InsertCaption(Anchor(session, "tbl"), new CaptionOptions { Label = "Table" }).Success);

        var blocks = Body(session).Elements().Where(e => e.Name != W + "sectPr").ToList();
        Assert.Equal(W + "p", blocks[1].Name);
        Assert.Equal("Table 1", Text(blocks[1]));
        Assert.NotNull(blocks[1].Element(W + "pPr")!.Element(W + "keepNext"));
        Assert.Equal(W + "tbl", blocks[2].Name);
    }

    [Fact]
    public void CP003_ACaptionBeforeOthers_NumbersThemAgain()
    {
        using var session = Open(P(R("Picture")) + Caption("Figure", "1", ": Old") + Caption("Table", "1"));

        Assert.True(session.InsertCaption(BodyParagraphs(session)[0], new CaptionOptions { Position = CaptionPosition.Above }).Success);

        Assert.Equal(new[] { "Figure 1", "Picture", "Figure 2: Old", "Table 1" }, Lines(session));
    }

    [Fact]
    public void CP004_ExcludedLabel_NumberFormat_AndALabelWithSpaces()
    {
        using var session = Open(P(R("One")) + P(R("Two")));
        var ps = BodyParagraphs(session);

        Assert.True(session.InsertCaption(ps[0], new CaptionOptions { Label = "My Plate", NumberFormat = "ROMAN" }).Success);
        Assert.True(session.InsertCaption(ps[1], new CaptionOptions { Label = "My Plate", NumberFormat = "ROMAN", ExcludeLabel = true, Text = ". Two" }).Success);

        Assert.Equal(new[] { "One", "My Plate I", "Two", "II. Two" }, Lines(session));
        Assert.All(Body(session).Descendants(W + "instrText"), i => Assert.Contains("SEQ My_Plate \\* ROMAN", i.Value));
    }

    [Theory]
    [InlineData("Fig\"ure", "ARABIC")]
    [InlineData("  ", "ARABIC")]
    [InlineData("Figure", "Ordinal")]
    public void CP005_ABadLabelOrFormat_IsRefused(string label, string format)
    {
        using var session = Open(P(R("One")));
        var before = Body(session).ToString();

        var result = session.InsertCaption(BodyParagraphs(session)[0], new CaptionOptions { Label = label, NumberFormat = format });

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.InvalidReferenceField, result.Error!.Code);
        Assert.Equal(before, Body(session).ToString());
    }

    [Fact]
    public void CP006_TheDocumentsOwnCaptionStyle_FoundByItsName()
    {
        using var session = Open(P(R("One")),
            "<w:style w:type=\"paragraph\" w:styleId=\"Beschriftung\"><w:name w:val=\"caption\"/></w:style>");

        Assert.True(session.InsertCaption(BodyParagraphs(session)[0]).Success);

        Assert.Equal("Beschriftung", StyleOf(Body(session).Elements(W + "p").ElementAt(1)));
        Assert.DoesNotContain(Styles(session).Elements(W + "style"), s => (string?)s.Attribute(W + "styleId") == "Caption");
    }

    [Fact]
    public void CP007_BelowTheParagraphThatEndsASection_TheSectionEndsAfterTheCaption()
    {
        using var session = Open(P(R("Picture"), null, "<w:sectPr/>") + P(R("Next section")));

        Assert.True(session.InsertCaption(BodyParagraphs(session)[0]).Success);

        var ps = Body(session).Elements(W + "p").ToList();
        Assert.Null(ps[0].Element(W + "pPr")?.Element(W + "sectPr"));
        Assert.NotNull(ps[1].Element(W + "pPr")!.Element(W + "sectPr"));
        Assert.Equal("Figure 1", Text(ps[1]));
    }

    [Fact]
    public void CP008_ListCaptions_GivesLabelNumberAndText_NotATablesEntries()
    {
        using var session = Open(Caption("Figure", "1", ": A cat") + P(R("Body")) + Caption("Table", "4", " Prices"));
        Assert.True(session.InsertTableOfFigures(BodyParagraphs(session)[0], Position.Before).Success);
        Assert.True(session.UpdateFields().Success);

        var captions = session.ListCaptions();

        Assert.Equal(new[] { ("Figure", "1", "Figure 1: A cat"), ("Table", "1", "Table 1 Prices") },
            captions.Select(c => (c.Label, c.Number, c.Text)).ToArray());
    }

    [Fact]
    public void CP009_ReferenceBookmarks_OverLabelAndNumber_OrTheTextAlone_Reused()
    {
        using var session = Open(Caption("Figure", "1", ": A cat"));
        var caption = BodyParagraphs(session)[0];

        var number = session.AddReferenceBookmark(caption, ReferenceTargetPart.LabelAndNumber);
        var text = session.AddReferenceBookmark(caption, ReferenceTargetPart.CaptionText);
        var whole = session.AddReferenceBookmark(caption, ReferenceTargetPart.Paragraph);

        Assert.True(number.Success, number.Error?.Message); Assert.True(text.Success, text.Error?.Message); Assert.True(whole.Success, whole.Error?.Message);
        var bookmarks = session.ListBookmarks().ToDictionary(b => b.Name, b => b.Text);
        Assert.Equal("Figure 1", bookmarks[number.BookmarkName!]);
        Assert.Equal("A cat", bookmarks[text.BookmarkName!]);
        Assert.Equal("Figure 1: A cat", bookmarks[whole.BookmarkName!]);
        Assert.Equal(number.BookmarkName, session.AddReferenceBookmark(caption, ReferenceTargetPart.LabelAndNumber).BookmarkName);
        Assert.Equal(text.BookmarkName, session.AddReferenceBookmark(caption, ReferenceTargetPart.CaptionText).BookmarkName);
        // (the label-and-number bookmark holds the whole field)
        var body = Body(session);
        var end = body.Descendants(W + "bookmarkEnd").Single(e => (string?)e.Attribute(W + "id") ==
            (string?)body.Descendants(W + "bookmarkStart").Single(s => (string?)s.Attribute(W + "name") == number.BookmarkName).Attribute(W + "id"));
        Assert.Equal("end", (string?)((XElement)end.PreviousNode!).Element(W + "fldChar")?.Attribute(W + "fldCharType"));
    }

    [Fact]
    public void CP010_ANotACaptionParagraph_HasNoLabelAndNumber()
    {
        using var session = Open(P(R("Plain")));

        var result = session.AddReferenceBookmark(BodyParagraphs(session)[0], ReferenceTargetPart.LabelAndNumber);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.AnchorWrongKind, result.Error!.Code);
    }

    [Fact]
    public void CP011_ATableOfFigures_IsFilledWithTheLabelsCaptions_AndLinked()
    {
        using var session = Open(P(R("Intro"), "Heading1") + Caption("Figure", "9", ": A cat") + Caption("Table", "1")
            + Caption("figure", "9", ": A dog"));
        var captions = session.ListCaptions();
        Assert.True(session.InsertTableOfFigures(BodyParagraphs(session)[0], Position.Before,
            new TableOfFiguresOptions { CaptionLabel = "Figure", Hyperlinks = true }).Success);

        var result = session.UpdateFields(new UpdateFieldsOptions
        {
            Pages = new Dictionary<string, string> { [captions[0].AnchorId] = "2", [captions[2].AnchorId] = "3" },
        });

        Assert.True(result.Success, result.Error?.Message);
        var lines = Lines(session);
        Assert.Equal(new[] { "Figure 1: A cat\t2", "figure 2: A dog\t3", "Intro", "Figure 1: A cat", "Table 1", "figure 2: A dog" }, lines);
        var entries = Body(session).Elements(W + "p").Take(2).ToList();
        Assert.All(entries, e => Assert.Equal("TableofFigures", StyleOf(e)));
        Assert.All(entries, e => Assert.NotNull(e.Element(W + "hyperlink")));
        Assert.Contains(Styles(session).Elements(W + "style"),
            s => (string?)s.Element(W + "name")?.Attribute(W + "val") == "table of figures");

        // Updated again, nothing changes: no undo step.
        Assert.True(session.UpdateFields(new UpdateFieldsOptions
        {
            Pages = new Dictionary<string, string> { [captions[0].AnchorId] = "2", [captions[2].AnchorId] = "3" },
        }).Success);
        Assert.Equal(lines, Lines(session));
    }

    [Fact]
    public void CP012_WithoutLabelAndNumber_AndWithNoCaptions()
    {
        using var session = Open(Contents(" TOC \\h \\z \\a \"Figure\" ") + Caption("Figure", "1", ": A cat")
            + Contents(" TOC \\h \\z \\c \"Plate\" "));

        Assert.True(session.UpdateFields().Success);

        Assert.Equal(new[] { "A cat\t1", "Figure 1: A cat", "No table of figures entries found." }, Lines(session));
    }

    private static string Contents(string instr) =>
        P(Start(instr) + R("old entry") + Char("end"), "TableofFigures");
}
