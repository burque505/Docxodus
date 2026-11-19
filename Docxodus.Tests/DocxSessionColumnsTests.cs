// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Xml.Linq;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Word's Layout ▸ Columns: <see cref="DocxSession.SetColumns"/> and its "Apply to" scopes.
/// </summary>
public class DocxSessionColumnsTests
{
    private const string Section =
        "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/>" +
        "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>" +
        "<w:cols w:space=\"720\"/></w:sectPr>";

    private const string ThreeParagraphs =
        "<w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p><w:p><w:r><w:t>three</w:t></w:r></w:p>" + Section;

    private static DocxSession Open(string body, TrackedChangeMode mode = TrackedChangeMode.Accept) =>
        DocxSessionClonedRevisionTests.Open(body, mode);

    private static XElement Body(DocxSession s) => DocxSessionClonedRevisionTests.Body(s);

    private static string[] Paragraphs(DocxSession s) =>
        Body(s).Elements(W.p).Select(p => string.Concat(p.Descendants(W.t).Select(t => t.Value))).ToArray();

    /// <summary>The body paragraphs' anchor ids, in document order.</summary>
    private static string[] Ids(DocxSession s) =>
        s.Project().AnchorIndex.Values.Select(v => v.Anchor)
            .Where(a => a.Kind == "p" && a.Scope == "body").Select(a => a.Id).ToArray();

    /// <summary>Each section's column count and start, in order (paragraph marks, then the body's).</summary>
    private static (string? Num, string? Type)[] Sections(DocxSession s) =>
        Body(s).Descendants(W.sectPr)
            .Select(x => ((string?)x.Element(W.cols)?.Attribute(W.num), (string?)x.Element(W.type)?.Attribute(W.val)))
            .ToArray();

    [Fact]
    public void CO001_Section_SetsTheCount_AndReadsBack()
    {
        using var session = Open(ThreeParagraphs);
        var id = Ids(session)[0];

        var res = session.SetColumns(id, 0, new ColumnsOp { Count = 2 });

        Assert.True(res.Success, res.Error?.Message);
        var cols = Body(session).Element(W.sectPr)!.Element(W.cols)!;
        Assert.Equal("2", (string?)cols.Attribute(W.num));
        Assert.Equal("720", (string?)cols.Attribute(W.space));
        var info = session.GetSectionInfo(id)!;
        Assert.Equal(2, info.Columns);
        Assert.Equal(720, info.ColumnSpaceTwips);
        Assert.False(info.ColumnSeparator);
        Assert.True(info.ColumnsEqualWidth);

        // The same again changes nothing and takes no undo step.
        Assert.True(session.SetColumns(id, 0, new ColumnsOp { Count = 2 }).Success);
        Assert.True(session.Undo());
        Assert.Null(Body(session).Element(W.sectPr)!.Element(W.cols)!.Attribute(W.num));
        Assert.False(session.Undo());
    }

    [Fact]
    public void CO002_SpaceAndLineBetween_AndOneColumnDropsTheCount()
    {
        using var session = Open(ThreeParagraphs);
        var id = Ids(session)[0];

        Assert.True(session.SetColumns(id, 0, new ColumnsOp { Count = 3, SpaceTwips = 360, Separator = true }).Success);
        var info = session.GetSectionInfo(id)!;
        Assert.Equal((3, 360, true), (info.Columns, info.ColumnSpaceTwips, info.ColumnSeparator));

        Assert.True(session.SetColumns(id, 0, new ColumnsOp { Count = 1, Separator = false }).Success);
        var cols = Body(session).Element(W.sectPr)!.Element(W.cols)!;
        Assert.Null(cols.Attribute(W.num));
        Assert.Null(cols.Attribute(W.sep));
        Assert.Equal(1, session.GetSectionInfo(id)!.Columns);
    }

    [Fact]
    public void CO003_ACount_MakesUnequalColumnsEqual()
    {
        const string unequal =
            "<w:p><w:r><w:t>text</w:t></w:r></w:p><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/>" +
            "<w:cols w:num=\"2\" w:equalWidth=\"0\" w:space=\"720\"><w:col w:w=\"2880\" w:space=\"720\"/><w:col w:w=\"5760\"/></w:cols></w:sectPr>";
        using var session = Open(unequal);
        var id = Ids(session)[0];
        Assert.False(session.GetSectionInfo(id)!.ColumnsEqualWidth);

        Assert.True(session.SetColumns(id, 0, new ColumnsOp { Count = 3 }).Success);

        var cols = Body(session).Element(W.sectPr)!.Element(W.cols)!;
        Assert.Empty(cols.Elements(W.col));
        Assert.Null(cols.Attribute(W.equalWidth));
        Assert.Equal("3", (string?)cols.Attribute(W.num));
        Assert.True(session.GetSectionInfo(id)!.ColumnsEqualWidth);
    }

    [Fact]
    public void CO004_Refusals_ChangeNothing()
    {
        using var session = Open(ThreeParagraphs);
        var id = Ids(session)[0];

        // 45 columns half an inch apart cannot fit 6.5 inches of text.
        var narrow = session.SetColumns(id, 0, new ColumnsOp { Count = 45 });
        Assert.False(narrow.Success);
        Assert.Equal(EditErrorCode.InvalidPageSetup, narrow.Error!.Code);
        Assert.Equal(EditErrorCode.InvalidPageSetup, session.SetColumns(id, 0, new ColumnsOp { Count = 0 }).Error!.Code);
        Assert.Equal(EditErrorCode.InvalidPageSetup, session.SetColumns(id, 0, new ColumnsOp { SpaceTwips = -1 }).Error!.Code);
        Assert.Equal(EditErrorCode.InvalidPageSetup, session.SetColumns(id, 0, new ColumnsOp { Count = 2 }, "sideways").Error!.Code);
        // Too narrow for the selection's section too: and no break is left behind.
        Assert.False(session.SetColumns(id, 1, new ColumnsOp { Count = 45 }, "selection", Ids(session)[1], 2).Success);

        Assert.Single(Body(session).Descendants(W.sectPr));
        Assert.Equal(new[] { "one", "two", "three" }, Paragraphs(session));
        Assert.False(session.Undo());
    }

    [Fact]
    public void CO005_Document_SetsEverySection_InOneUndoStep()
    {
        const string twoSections =
            "<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:pPr><w:r><w:t>one</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>two</w:t></w:r></w:p>" + Section;
        using var session = Open(twoSections);

        Assert.True(session.SetColumns(Ids(session)[1], 0, new ColumnsOp { Count = 2 }, "document").Success);
        Assert.Equal(new (string?, string?)[] { ("2", null), ("2", null) }, Sections(session));

        Assert.True(session.Undo());
        Assert.All(Sections(session), s => Assert.Null(s.Num));
        Assert.False(session.Undo());
    }

    [Fact]
    public void CO006_Forward_InAParagraph_SplitsIt_AndColumnsStartAtTheCaret()
    {
        using var session = Open(ThreeParagraphs);

        var res = session.SetColumns(Ids(session)[1], 1, new ColumnsOp { Count = 2 }, "forward");

        Assert.True(res.Success, res.Error?.Message);
        Assert.Equal(new[] { "one", "t", "wo", "three" }, Paragraphs(session));
        // The text before keeps one column; the new section starts continuously with two.
        Assert.Equal(new[] { ((string?)null, (string?)null), ("2", "continuous") }, Sections(session));
        Assert.NotNull(Body(session).Elements(W.p).ElementAt(1).Element(W.pPr)?.Element(W.sectPr));

        Assert.True(session.Undo());
        Assert.Equal(new[] { "one", "two", "three" }, Paragraphs(session));
        Assert.Single(Body(session).Descendants(W.sectPr));
        Assert.False(session.Undo());
    }

    [Fact]
    public void CO007_Forward_AtTheStartOfAParagraph_BreaksOnThePreviousMark()
    {
        using var session = Open(ThreeParagraphs);

        Assert.True(session.SetColumns(Ids(session)[1], 0, new ColumnsOp { Count = 2 }, "forward").Success);

        Assert.Equal(new[] { "one", "two", "three" }, Paragraphs(session));
        Assert.NotNull(Body(session).Elements(W.p).First().Element(W.pPr)?.Element(W.sectPr));
        Assert.Equal(new[] { ((string?)null, (string?)null), ("2", "continuous") }, Sections(session));

        // At the start of the document there is no section to end.
        using var atStart = Open(ThreeParagraphs);
        Assert.True(atStart.SetColumns(Ids(atStart)[0], 0, new ColumnsOp { Count = 2 }, "forward").Success);
        Assert.Equal(new (string?, string?)[] { ("2", null) }, Sections(atStart));
    }

    [Fact]
    public void CO008_Selection_OfWholeParagraphs_PutsBreaksOnTheirMarks()
    {
        using var session = Open(ThreeParagraphs);
        var ids = Ids(session);

        var res = session.SetColumns(ids[1], 0, new ColumnsOp { Count = 2 }, "selection", ids[1], 3);

        Assert.True(res.Success, res.Error?.Message);
        Assert.Equal(new[] { "one", "two", "three" }, Paragraphs(session));
        // one | two in two columns, starting continuously | three, one column, continuous again.
        Assert.Equal(new[] { ((string?)null, (string?)null), ("2", "continuous"), (null, "continuous") }, Sections(session));
        Assert.Equal(2, session.GetSectionInfo(ids[1])!.Columns);
        Assert.Equal(1, session.GetSectionInfo(ids[2])!.Columns);

        Assert.True(session.Undo());
        Assert.Single(Body(session).Descendants(W.sectPr));
        Assert.False(session.Undo());
    }

    [Fact]
    public void CO009_Selection_InsideOneParagraph_SplitsItTwice()
    {
        using var session = Open("<w:p><w:r><w:t>abcdef</w:t></w:r></w:p>" + Section);
        var id = Ids(session)[0];

        var res = session.SetColumns(id, 2, new ColumnsOp { Count = 3 }, "selection", id, 4);

        Assert.True(res.Success, res.Error?.Message);
        Assert.Equal(new[] { "ab", "cd", "ef" }, Paragraphs(session));
        Assert.Equal(new[] { ((string?)null, (string?)null), ("3", "continuous"), (null, "continuous") }, Sections(session));
    }

    [Fact]
    public void CO010_Selection_ToTheEndOfTheDocument_KeepsALastParagraph()
    {
        using var session = Open(ThreeParagraphs);
        var ids = Ids(session);

        Assert.True(session.SetColumns(ids[1], 0, new ColumnsOp { Count = 2 }, "selection", ids[2], 5).Success);

        Assert.Equal(new[] { "one", "two", "three", "" }, Paragraphs(session));
        Assert.Equal(new[] { ((string?)null, (string?)null), ("2", "continuous"), (null, "continuous") }, Sections(session));
    }

    [Fact]
    public void CO011_Selection_EndingInATable_BreaksAfterTheTable()
    {
        const string withTable =
            "<w:p><w:r><w:t>one</w:t></w:r></w:p>" +
            "<w:tbl><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
            "<w:p><w:r><w:t>after</w:t></w:r></w:p>" + Section;
        using var session = Open(withTable);
        var ids = Ids(session);
        var cell = ids[1];

        Assert.True(session.SetColumns(ids[0], 0, new ColumnsOp { Count = 2 }, "selection", cell, 2).Success);

        var body = Body(session);
        var afterTable = body.Element(W.tbl)!.ElementsAfterSelf().First();
        Assert.Equal(W.p, afterTable.Name);
        Assert.Equal("", string.Concat(afterTable.Descendants(W.t).Select(t => t.Value)));
        Assert.Equal("2", (string?)afterTable.Element(W.pPr)!.Element(W.sectPr)!.Element(W.cols)!.Attribute(W.num));
        Assert.Equal(new (string?, string?)[] { ("2", null), (null, "continuous") }, Sections(session));
    }

    [Fact]
    public void CO012_Tracked_SectionIsAPropertyChange_ButNewSectionsAreRefused()
    {
        using var session = Open(ThreeParagraphs, TrackedChangeMode.RenderInline);
        var id = Ids(session)[0];

        var forward = session.SetColumns(id, 1, new ColumnsOp { Count = 2 }, "forward");
        Assert.False(forward.Success);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported, forward.Error!.Code);
        Assert.Single(Body(session).Descendants(W.sectPr));

        Assert.True(session.SetColumns(id, 0, new ColumnsOp { Count = 2 }).Success);
        Assert.NotNull(Body(session).Element(W.sectPr)!.Element(W.sectPrChange));
    }

    [Fact]
    public void CO013_TheLineBetween_IsDrawnInTheHtml()
    {
        using var session = Open(ThreeParagraphs);
        Assert.True(session.SetColumns(Ids(session)[0], 0, new ColumnsOp { Count = 2, Separator = true }).Success);

        var wml = new WmlDocument("columns.docx", session.Save());
        using var stream = new OpenXmlMemoryStreamDocument(wml);
        using var doc = stream.GetWordprocessingDocument();
        var html = WmlToHtmlConverter.ConvertToHtml(doc, new WmlToHtmlConverterSettings());

        var section = html.Descendants().Single(e => (string?)e.Attribute("data-cols") == "2");
        Assert.Equal("true", (string?)section.Attribute("data-col-sep"));
        Assert.Contains("column-rule: 0.75pt solid currentColor;", (string?)section.Attribute("style"));
    }
}
