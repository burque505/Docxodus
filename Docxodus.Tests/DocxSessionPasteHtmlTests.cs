// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>DocxSession.InsertHtml: pasting the clipboard's HTML at a caret.</summary>
public class DocxSessionPasteHtmlTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const string Rich = "<html><body><h2>Title here</h2><p style=\"font-family: Arial; font-size: 20pt; color: red\">Some <b>bold</b> and "
        + "<a href=\"https://example.com/\">a link</a>.</p><ul><li>one</li><li>two<ul><li>inner</li></ul></li></ul>"
        + "<ol><li>first</li></ol><table><tr><td>a</td><td>b</td></tr></table></body></html>";

    /// <summary>"First paragraph." in Georgia 14 pt (direct formatting), "Second paragraph." plain.</summary>
    private static byte[] Fixture()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(
                    new RunProperties(new RunFonts { Ascii = "Georgia", HighAnsi = "Georgia" }, new FontSize { Val = "28" }),
                    new Text("First paragraph."))),
                new Paragraph(new Run(new Text("Second paragraph.")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = DocxSessionTests.BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static string[] Blocks(DocxSession session) =>
        session.Project().AnchorIndex.Values.Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h" or "li")
            .Select(t => t.Anchor.Id).Distinct().ToArray();

    private static XElement Body(DocxSession session)
    {
        using var stream = new MemoryStream(session.Save(false));
        using var doc = WordprocessingDocument.Open(stream, false);
        return XDocument.Load(doc.MainDocumentPart!.GetStream()).Root!.Element(W + "body")!;
    }

    private static string Text(XElement p) =>
        string.Concat(p.Descendants().Where(e => e.Name == W + "t").Select(e => e.Value));

    private static XElement[] Paragraphs(DocxSession session) =>
        Body(session).Elements(W + "p").ToArray();

    [Fact]
    public void PH001_SinglePlainParagraph_GoesIntoTheText()
    {
        using var session = new DocxSession(Fixture());
        var first = Blocks(session)[0];
        var result = session.InsertHtml(first, 6, 0, "<html><body><p>pasted <b>bold</b>, </p></body></html>");
        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal("First pasted bold,paragraph.", Text(Paragraphs(session)[0]));
        Assert.Equal(2, Paragraphs(session).Length);
        Assert.Equal(6 + "pasted bold,".Length, result.Caret!.Offset);
        Assert.Equal(first, result.Caret.AnchorId);
    }

    [Fact]
    public void PH002_SeveralBlocks_SplitTheParagraph_FirstAndLastJoinIt()
    {
        using var session = new DocxSession(Fixture());
        var result = session.InsertHtml(Blocks(session)[0], 6, 0,
            "<html><body><p>one</p><h2>Two</h2><p>three</p></body></html>");
        Assert.True(result.Success, result.Error?.Message);
        var texts = Paragraphs(session).Select(Text).ToArray();
        Assert.Equal(new[] { "First one", "Two", "threeparagraph.", "Second paragraph." }, texts);
        Assert.Equal("Heading2", (string?)Paragraphs(session)[1].Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val"));
        Assert.Equal("three".Length, result.Caret!.Offset);
    }

    [Fact]
    public void PH003_ListsTablesAndLinks_Arrive_WithTheirOwnNumbering()
    {
        using var session = new DocxSession(Fixture());
        var result = session.InsertHtml(Blocks(session)[1], 0, 0, Rich);
        Assert.True(result.Success, result.Error?.Message);
        var body = Body(session);
        Assert.Single(body.Elements(W + "tbl"));
        var numIds = body.Descendants(W + "numId").Select(n => (string)n.Attribute(W + "val")!).Distinct().ToArray();
        Assert.Equal(2, numIds.Length);   // the bulleted list and the numbered one

        using var stream = new MemoryStream(session.Save(false));
        using var doc = WordprocessingDocument.Open(stream, false);
        var numbering = XDocument.Load(doc.MainDocumentPart!.NumberingDefinitionsPart!.GetStream()).Root!;
        foreach (var id in numIds)
            Assert.Single(numbering.Elements(W + "num"), n => (string?)n.Attribute(W + "numId") == id);
        var link = Assert.Single(body.Descendants(W + "hyperlink"));
        var rel = doc.MainDocumentPart.HyperlinkRelationships.Single(r => r.Id == (string?)link.Attribute(R + "id"));
        Assert.Equal("https://example.com/", rel.Uri.OriginalString);
        // Pasted at the start of "Second paragraph.": it stays, after the table.
        Assert.Equal("Second paragraph.", Text(body.Elements(W + "p").Last()));
    }

    [Fact]
    public void PH004_Merge_DropsTheSourceLook_AndTakesTheLandingRunsLook()
    {
        using var session = new DocxSession(Fixture());
        var result = session.InsertHtml(Blocks(session)[0], 16, 0,
            "<html><body><p style=\"font-family: Arial; font-size: 20pt; color: red\">A <b>B</b></p><p>C</p></body></html>");
        Assert.True(result.Success, result.Error?.Message);
        var runs = Paragraphs(session).Take(2).SelectMany(p => p.Descendants(W + "r")).ToArray();
        Assert.All(runs, r => Assert.Equal("Georgia", (string?)r.Element(W + "rPr")?.Element(W + "rFonts")?.Attribute(W + "ascii")));
        Assert.All(runs, r => Assert.Equal("28", (string?)r.Element(W + "rPr")?.Element(W + "sz")?.Attribute(W + "val")));
        Assert.DoesNotContain(runs, r => r.Element(W + "rPr")?.Element(W + "color") is not null);
        Assert.Contains(runs, r => Text(r) == "B" && r.Element(W + "rPr")?.Element(W + "b") is not null);
    }

    [Fact]
    public void PH005_Keep_KeepsTheSourceLook()
    {
        using var session = new DocxSession(Fixture());
        var result = session.InsertHtml(Blocks(session)[1], 0, 0,
            "<html><body><p style=\"font-family: Arial; color: red\">A</p></body></html>",
            new HtmlInsertOptions { Formatting = HtmlPasteFormatting.Keep });
        Assert.True(result.Success, result.Error?.Message);
        var run = Paragraphs(session)[1].Descendants(W + "r").First(r => Text(r) == "A");
        Assert.Equal("FF0000", (string?)run.Element(W + "rPr")?.Element(W + "color")?.Attribute(W + "val"), ignoreCase: true);
    }

    [Fact]
    public void PH006_OneUndoStep_EvenOverASelection()
    {
        using var session = new DocxSession(Fixture());
        var before = session.Project().Markdown;
        var result = session.InsertHtml(Blocks(session)[0], 0, 5, Rich);
        Assert.True(result.Success, result.Error?.Message);
        Assert.DoesNotContain("First", Text(Body(session)));
        Assert.True(session.Undo());
        Assert.Equal(before, session.Project().Markdown);
        Assert.False(session.Undo());
    }

    [Fact]
    public void PH007_Pictures_GetPartsAndFreeIds()
    {
        using var session = new DocxSession(Fixture());
        Assert.True(session.InsertImage(Blocks(session)[1], 0, DocxSessionImageTests.Png(2, 2)).Success);
        const string png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";   // a real 1x1 PNG: the converter decodes it
        var result = session.InsertHtml(Blocks(session)[0], 0, 0,
            $"<html><body><p>x<img src=\"data:image/png;base64,{png}\"/></p><p>y</p></body></html>");
        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal(2, session.ListImages().Count);
        var ids = Body(session).Descendants(WP + "docPr").Select(d => (string)d.Attribute("id")!).ToArray();
        Assert.Equal(ids.Length, ids.Distinct().Count());
        PackageValidation.AssertValid(session.Save(true), DocumentFormat.OpenXml.FileFormatVersions.Office2007,
            e => !e.Description.Contains("powertools.codeplex.com", StringComparison.Ordinal));
    }

    [Fact]
    public void PH008_EmptyParagraph_BecomesTheLastPastedParagraph()
    {
        using var session = new DocxSession(Fixture());
        var created = session.InsertParagraph(Blocks(session)[0], Position.After, "x");
        var empty = created.Created[0].Id;
        Assert.True(session.ReplaceTextAtSpan(empty, 0, 1, "").Success);
        var result = session.InsertHtml(Blocks(session)[1], 0, 0, "<html><body><p>one</p><h2>Two</h2></body></html>");
        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal(new[] { "First paragraph.", "one", "Two", "Second paragraph." }, Paragraphs(session).Select(Text).ToArray());
    }

    [Fact]
    public void PH009_Tracked_IsAnInsertion()
    {
        using var session = new DocxSession(Fixture());
        session.SetTrackedChanges(TrackedChangeMode.RenderInline);
        var result = session.InsertHtml(Blocks(session)[0], 6, 0, Rich);
        Assert.True(result.Success, result.Error?.Message);
        var body = Body(session);
        Assert.All(body.Descendants(W + "r").Where(r => Text(r) is "Title here" or "one" or "a"),
            r => Assert.Contains(r.Ancestors(), a => a.Name == W + "ins"));
        Assert.All(body.Elements(W + "tbl").Descendants(W + "trPr"), trPr => Assert.NotNull(trPr.Element(W + "ins")));
        PackageValidation.AssertValid(session.Save(true), DocumentFormat.OpenXml.FileFormatVersions.Office2007,
            e => !e.Description.Contains("powertools.codeplex.com", StringComparison.Ordinal));
    }

    [Fact]
    public void PH011_PlainText_LinesAreParagraphs_TabsAreTabs_OneUndoStep()
    {
        using var session = new DocxSession(Fixture());
        var before = session.Project().Markdown;
        var result = session.InsertPastedText(Blocks(session)[0], 6, 0, "one\ttwo\r\nthree\n");
        Assert.True(result.Success, result.Error?.Message);
        var paragraphs = Paragraphs(session);
        Assert.Equal(new[] { "First onetwo", "three", "paragraph.", "Second paragraph." }, paragraphs.Select(Text).ToArray());
        Assert.Single(paragraphs[0].Descendants(W + "tab"));
        // The pasted text takes the look where it lands (Georgia 14 pt).
        Assert.All(paragraphs[1].Descendants(W + "r"),
            r => Assert.Equal("Georgia", (string?)r.Element(W + "rPr")?.Element(W + "rFonts")?.Attribute(W + "ascii")));
        Assert.Equal(0, result.Caret!.Offset);   // the caret ends before "paragraph."
        Assert.True(session.Undo());
        Assert.Equal(before, session.Project().Markdown);
    }

    [Fact]
    public void PH012_PlainText_OneLine_GoesIntoTheText()
    {
        using var session = new DocxSession(Fixture());
        var result = session.InsertPastedText(Blocks(session)[1], 0, 6, "Other");
        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal("Other paragraph.", Text(Paragraphs(session)[1]));
        Assert.Equal(5, result.Caret!.Offset);
    }

    [Fact]
    public void PH013_BrowserCopyColors_RgbaAndFunctions_DoNotFailThePaste()
    {
        // A browser's copy writes computed styles on every element: transparent rgba() backgrounds, rgb() colors.
        using var session = new DocxSession(Fixture());
        var result = session.InsertHtml(Blocks(session)[1], 0, 0,
            "<html><body><p style=\"color: rgb(0, 0, 0); background-color: rgba(0, 0, 0, 0)\">A <span style=\"color: rgba(255, 0, 0, 0.5); "
            + "background-color: hsl(0, 0%, 50%)\">B</span></p><p>C</p></body></html>",
            new HtmlInsertOptions { Formatting = HtmlPasteFormatting.Keep });
        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal("A B", Text(Paragraphs(session)[1]));
        var b = Paragraphs(session)[1].Descendants(W + "r").First(r => Text(r) == "B");
        Assert.Equal("ff0000", (string?)b.Element(W + "rPr")?.Element(W + "color")?.Attribute(W + "val"), ignoreCase: true);
    }

    [Fact]
    public void PH010_NotXml_IsRefused_AndChangesNothing()
    {
        using var session = new DocxSession(Fixture());
        var before = session.Project().Markdown;
        var result = session.InsertHtml(Blocks(session)[0], 0, 0, "<p>unclosed");
        Assert.False(result.Success);
        Assert.Equal(before, session.Project().Markdown);
        Assert.False(session.Undo());
    }
}
