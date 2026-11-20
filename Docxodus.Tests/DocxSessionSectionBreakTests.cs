// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Xml.Linq;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Word's Layout ▸ Breaks: <see cref="DocxSession.InsertSectionBreak"/>,
/// <see cref="DocxSession.InsertBreak"/> and the section start of <see cref="DocxSession.SetPageSetup"/>.
/// </summary>
public class DocxSessionSectionBreakTests
{
    private const string Section =
        "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/>" +
        "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>" +
        "<w:pgNumType w:fmt=\"lowerRoman\" w:start=\"1\"/><w:cols w:space=\"720\"/></w:sectPr>";

    private const string OneParagraph = "<w:p><w:r><w:t>abcdef</w:t></w:r></w:p>" + Section;

    private static DocxSession Open(string body, TrackedChangeMode mode = TrackedChangeMode.Accept) =>
        DocxSessionClonedRevisionTests.Open(body, mode);

    private static string First(DocxSession s) => DocxSessionClonedRevisionTests.FirstParagraph(s);

    private static XElement Body(DocxSession s) => DocxSessionClonedRevisionTests.Body(s);

    private static string[] Paragraphs(DocxSession s) =>
        Body(s).Elements(W.p).Select(p => string.Concat(p.Descendants(W.t).Select(t => t.Value))).ToArray();

    [Fact]
    public void SB001_NextPage_SplitsTheParagraph_AndTheFirstHalfEndsASectionLikeTheOld()
    {
        using var session = Open(OneParagraph);

        var res = session.InsertSectionBreak(First(session), 3, "nextPage");

        Assert.True(res.Success, res.Error?.Message);
        Assert.Equal(new[] { "abc", "def" }, Paragraphs(session));
        var body = Body(session);
        var ending = body.Elements(W.p).First().Element(W.pPr)!.Element(W.sectPr)!;
        var final = body.Element(W.sectPr)!;
        Assert.Equal("12240", (string?)ending.Element(W.pgSz)!.Attribute(W._w));
        Assert.Equal("1440", (string?)ending.Element(W.pgMar)!.Attribute(W.left));
        // The restart stays with the first section; the new one numbers on in the same format.
        Assert.Equal("1", (string?)ending.Element(W.pgNumType)!.Attribute(W.start));
        Assert.Null(final.Element(W.pgNumType)!.Attribute(W.start));
        Assert.Equal("lowerRoman", (string?)final.Element(W.pgNumType)!.Attribute(W.fmt));
        Assert.Null(final.Element(W.type));
        Assert.Null(body.Elements(W.p).Last().Element(W.pPr)?.Element(W.sectPr));

        var created = res.Created.Single().Id;
        Assert.Equal("nextPage", session.GetSectionInfo(created)!.SectionStart);
        Assert.Null(session.GetSectionInfo(created)!.PageNumberStart);
        Assert.Equal(1, session.GetSectionInfo(First(session))!.PageNumberStart);
    }

    [Theory]
    [InlineData("continuous")]
    [InlineData("evenPage")]
    [InlineData("oddPage")]
    [InlineData("nextColumn")]
    public void SB002_TheStartIsWrittenOnTheSectionAfterTheBreak(string start)
    {
        using var session = Open(OneParagraph);

        var res = session.InsertSectionBreak(First(session), 6, start);

        Assert.True(res.Success, res.Error?.Message);
        Assert.Equal(new[] { "abcdef", "" }, Paragraphs(session));
        var body = Body(session);
        Assert.Equal(start, (string?)body.Element(W.sectPr)!.Element(W.type)!.Attribute(W.val));
        Assert.Null(body.Elements(W.p).First().Element(W.pPr)!.Element(W.sectPr)!.Element(W.type));
        Assert.Equal(start, session.GetSectionInfo(res.Created.Single().Id)!.SectionStart);
        Assert.Equal("nextPage", session.GetSectionInfo(First(session))!.SectionStart);
    }

    [Fact]
    public void SB003_InAParagraphThatAlreadyEndsASection_TheOldBreakStaysOnTheSecondHalf()
    {
        const string twoSections =
            "<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:type w:val=\"continuous\"/></w:sectPr></w:pPr>" +
            "<w:r><w:t>abcdef</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>last</w:t></w:r></w:p>" +
            "<w:sectPr><w:pgSz w:w=\"15840\" w:h=\"12240\" w:orient=\"landscape\"/></w:sectPr>";
        using var session = Open(twoSections);

        var res = session.InsertSectionBreak(First(session), 2, "oddPage");

        Assert.True(res.Success, res.Error?.Message);
        var paragraphs = Body(session).Elements(W.p).ToList();
        Assert.Equal(3, paragraphs.Count);
        var firstEnd = paragraphs[0].Element(W.pPr)!.Element(W.sectPr)!;
        var secondEnd = paragraphs[1].Element(W.pPr)!.Element(W.sectPr)!;
        Assert.Equal("continuous", (string?)firstEnd.Element(W.type)!.Attribute(W.val));
        Assert.Equal("oddPage", (string?)secondEnd.Element(W.type)!.Attribute(W.val));
        Assert.Null(paragraphs[2].Element(W.pPr)?.Element(W.sectPr));
        Assert.Equal("landscape", (string?)Body(session).Element(W.sectPr)!.Element(W.pgSz)!.Attribute(W.orient));
    }

    [Fact]
    public void SB004_IsOneUndoStep_AndSurvivesSaving()
    {
        using var session = Open(OneParagraph);

        var res = session.InsertSectionBreak(First(session), 3, "continuous");
        Assert.True(res.Success, res.Error?.Message);
        using (var reopened = new DocxSession(session.Save()))
        {
            Assert.Equal(new[] { "abc", "def" }, Paragraphs(reopened));
            var second = reopened.Project().AnchorIndex.Values.Select(v => v.Anchor).Where(a => a.Kind == "p").Last().Id;
            Assert.Equal("continuous", reopened.GetSectionInfo(second)!.SectionStart);
        }

        Assert.True(session.Undo());
        Assert.Equal(new[] { "abcdef" }, Paragraphs(session));
        Assert.Single(Body(session).Descendants(W.sectPr));
        Assert.Equal("1", (string?)Body(session).Element(W.sectPr)!.Element(W.pgNumType)!.Attribute(W.start));
        Assert.True(session.Redo());
        Assert.Equal(new[] { "abc", "def" }, Paragraphs(session));
    }

    [Fact]
    public void SB005_Refusals_ChangeNothing()
    {
        const string inTable =
            "<w:tbl><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
            OneParagraph;
        using (var session = Open(inTable))
        {
            // The first paragraph anchor in document order is the one in the cell.
            var cell = First(session);
            var res = session.InsertSectionBreak(cell, 2, "nextPage");
            Assert.False(res.Success);
            Assert.Equal(EditErrorCode.AnchorWrongKind, res.Error!.Code);
            var br = session.InsertBreak(cell, 2, "page");
            Assert.False(br.Success);
        }
        using (var session = Open(OneParagraph))
        {
            var res = session.InsertSectionBreak(First(session), 3, "sideways");
            Assert.False(res.Success);
            Assert.Equal(EditErrorCode.InvalidPageSetup, res.Error!.Code);
            var br = session.InsertBreak(First(session), 3, "line");
            Assert.False(br.Success);
            var range = session.InsertSectionBreak(First(session), 99, "nextPage");
            Assert.False(range.Success);
            Assert.Equal(new[] { "abcdef" }, Paragraphs(session));
            Assert.False(session.Undo());
        }
        using (var session = Open(OneParagraph, TrackedChangeMode.RenderInline))
        {
            var res = session.InsertSectionBreak(First(session), 3, "nextPage");
            Assert.False(res.Success);
            Assert.Equal(EditErrorCode.TrackedOperationUnsupported, res.Error!.Code);
            Assert.Equal(new[] { "abcdef" }, Paragraphs(session));
        }
    }

    [Theory]
    [InlineData("page", 3, "abc", "def")]
    [InlineData("column", 3, "abc", "def")]
    [InlineData("page", 0, "", "abcdef")]
    [InlineData("page", 6, "abcdef", "")]
    public void SB006_InsertBreak_EndsTheFirstHalfWithTheBreak(string kind, int offset, string before, string after)
    {
        const string bold = "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>abcdef</w:t></w:r></w:p>" + Section;
        using var session = Open(bold);

        var res = session.InsertBreak(First(session), offset, kind);

        Assert.True(res.Success, res.Error?.Message);
        Assert.Equal(new[] { before, after }, Paragraphs(session));
        var first = Body(session).Elements(W.p).First();
        var lastRun = first.Elements(W.r).Last();
        Assert.Equal(kind, (string?)lastRun.Element(W.br)!.Attribute(W.type));
        if (offset > 0) Assert.NotNull(lastRun.Element(W.rPr)?.Element(W.b));
        Assert.Single(Body(session).Descendants(W.sectPr));
        Assert.True(session.Undo());
        Assert.Equal(new[] { "abcdef" }, Paragraphs(session));
        Assert.Empty(Body(session).Descendants(W.br));
    }

    [Fact]
    public void SB007_PageSetup_SetsAndClearsTheSectionStart()
    {
        using var session = Open(OneParagraph);
        var anchor = First(session);

        Assert.True(session.SetPageSetup(anchor, new PageSetupOp { SectionStart = "evenPage" }).Success);
        Assert.Equal("evenPage", session.GetSectionInfo(anchor)!.SectionStart);
        Assert.Equal("evenPage", (string?)Body(session).Element(W.sectPr)!.Element(W.type)!.Attribute(W.val));

        Assert.True(session.SetPageSetup(anchor, new PageSetupOp { SectionStart = "nextPage" }).Success);
        Assert.Null(Body(session).Element(W.sectPr)!.Element(W.type));
        Assert.Equal("nextPage", session.GetSectionInfo(anchor)!.SectionStart);

        var bad = session.SetPageSetup(anchor, new PageSetupOp { SectionStart = "nope" });
        Assert.False(bad.Success);
        Assert.Equal(EditErrorCode.InvalidPageSetup, bad.Error!.Code);

        // The wire form carries it too.
        var json = Docxodus.Internal.DocxSessionJson.ParsePageSetupOp("{\"sectionStart\":\"continuous\"}");
        Assert.Equal("continuous", json.SectionStart);
        Assert.Contains("\"sectionStart\":\"nextPage\"",
            Docxodus.Internal.DocxSessionJson.SerializeSectionInfoOrNull(session.GetSectionInfo(anchor)));
    }

    [Fact]
    public void SB008_ADocumentWithoutSectionProperties_GetsBoth()
    {
        using var session = Open("<w:p><w:r><w:t>abcdef</w:t></w:r></w:p>");

        var res = session.InsertSectionBreak(First(session), 3, "continuous");

        Assert.True(res.Success, res.Error?.Message);
        var body = Body(session);
        Assert.NotNull(body.Elements(W.p).First().Element(W.pPr)!.Element(W.sectPr));
        Assert.Equal("continuous", (string?)body.Element(W.sectPr)!.Element(W.type)!.Attribute(W.val));
    }

    private static string Columns(int n) =>
        $"<w:sectPr><w:type w:val=\"continuous\"/><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:cols w:num=\"{n}\" w:space=\"720\"/></w:sectPr>";

    [Fact]
    public void SB009_RemoveSectionBreak_TheTextBeforeJoinsTheSectionAfter_OneUndoStep()
    {
        // abc ends a one-column section; def is in a three-column section.
        using var session = Open("<w:p><w:pPr>" + Columns(1) + "</w:pPr><w:r><w:t>abc</w:t></w:r></w:p>"
            + "<w:p><w:r><w:t>def</w:t></w:r></w:p>" + Columns(3));
        var first = First(session);
        Assert.Equal(1, session.GetSectionInfo(first)!.Columns);

        var res = session.RemoveSectionBreak(first);

        Assert.True(res.Success, res.Error?.Message);
        Assert.Equal(new[] { "abc", "def" }, Paragraphs(session));
        Assert.Empty(Body(session).Elements(W.p).Where(p => p.Element(W.pPr)?.Element(W.sectPr) is not null));
        Assert.Equal(3, session.GetSectionInfo(First(session))!.Columns);

        Assert.True(session.Undo());
        Assert.Equal(1, session.GetSectionInfo(First(session))!.Columns);
        Assert.Single(Body(session).Elements(W.p).Where(p => p.Element(W.pPr)?.Element(W.sectPr) is not null));
    }

    [Fact]
    public void SB010_RemoveSectionBreak_AParagraphThatOnlyHeldTheBreakGoesWithIt()
    {
        using var session = Open("<w:p><w:r><w:t>abc</w:t></w:r></w:p><w:p><w:pPr>" + Columns(1) + "</w:pPr></w:p>"
            + "<w:p><w:r><w:t>def</w:t></w:r></w:p>" + Columns(2));
        var mark = session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p")
            .Select(t => t.Anchor.Id).ElementAt(1);

        var res = session.RemoveSectionBreak(mark);

        Assert.True(res.Success, res.Error?.Message);
        Assert.Equal(mark, res.Removed.Single().Id);
        Assert.Equal(new[] { "abc", "def" }, Paragraphs(session));
        Assert.Equal(2, session.GetSectionInfo(First(session))!.Columns);
        Assert.True(session.Undo());
        Assert.Equal(new[] { "abc", "", "def" }, Paragraphs(session));
    }

    [Fact]
    public void SB011_RemoveSectionBreak_Refusals_ChangeNothing()
    {
        using var plain = Open(OneParagraph);
        var noBreak = plain.RemoveSectionBreak(First(plain));
        Assert.False(noBreak.Success);
        Assert.Equal(EditErrorCode.AnchorWrongKind, noBreak.Error!.Code);

        using var tracked = Open("<w:p><w:pPr>" + Columns(1) + "</w:pPr><w:r><w:t>abc</w:t></w:r></w:p>"
            + "<w:p><w:r><w:t>def</w:t></w:r></w:p>" + Columns(3), TrackedChangeMode.RenderInline);
        long version = tracked.Version;
        var refused = tracked.RemoveSectionBreak(First(tracked));
        Assert.False(refused.Success);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported, refused.Error!.Code);
        Assert.Equal(version, tracked.Version);
    }
}
