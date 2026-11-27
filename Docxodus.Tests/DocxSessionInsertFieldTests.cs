using System.Linq;
using System.Xml.Linq;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary><see cref="DocxSession.InsertField"/>: field codes with nested fields in braces.</summary>
public class DocxSessionInsertFieldTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static string[] BodyParagraphs(DocxSession session) =>
        session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Kind == "p" && t.Anchor.Scope == "body")
            .Select(t => t.Anchor.Id).ToArray();

    private static XElement Para(DocxSession session, string id) => XElement.Parse(session.Raw.GetXml(id));

    private static string[] FldChars(XElement p) =>
        p.Descendants(W + "fldChar").Select(f => (string)f.Attribute(W + "fldCharType")!).ToArray();

    private static string Instr(XElement p) => string.Concat(p.Descendants(W + "instrText").Select(t => t.Value));

    [Fact]
    public void IF001_MergeField_ShowsGuillemetNameAtTheOffset()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var p = BodyParagraphs(session)[0];
        var before = session.GetAnchorInfo(p)!.VisibleText;

        var r = session.InsertField(p, 2, @"MERGEFIELD BoFull \* Upper");

        Assert.True(r.Success, r.Error?.Message);
        var xml = Para(session, p);
        Assert.Equal(new[] { "begin", "separate", "end" }, FldChars(xml));
        Assert.Equal(@"MERGEFIELD BoFull \* Upper", Instr(xml));
        Assert.Equal(before[..2] + "«BoFull»" + before[2..], session.GetAnchorInfo(p)!.VisibleText);
        Assert.True(session.Undo());
        Assert.Equal(before, session.GetAnchorInfo(p)!.VisibleText);
    }

    [Fact]
    public void IF002_NestedFields_WrittenInsideTheInstruction()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var p = BodyParagraphs(session)[0];

        var r = session.InsertField(p, 0, @"{ IF { MERGEFIELD IsGuaranty } = True ""Guaranty"" """" }");

        Assert.True(r.Success, r.Error?.Message);
        var xml = Para(session, p);
        Assert.Equal(new[] { "begin", "begin", "end", "separate", "end" }, FldChars(xml));
        Assert.Equal(@" IF  MERGEFIELD IsGuaranty  = True ""Guaranty"" """" ", Instr(xml));
        Assert.StartsWith("«IF IsGuaranty = True \"Guaranty\" \"\"»First", session.GetAnchorInfo(p)!.VisibleText);
    }

    [Fact]
    public void IF003_Formula_SummaryDropsSwitches_CachedResultWins()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var ps = BodyParagraphs(session);

        Assert.True(session.InsertField(ps[0], 0, @"= { MERGEFIELD Rate } * 100 \# ""0.00%""").Success);
        Assert.StartsWith("«= Rate * 100»", session.GetAnchorInfo(ps[0])!.VisibleText);

        Assert.True(session.InsertField(ps[1], 0, "FILLIN \"Name?\"", "[name]").Success);
        Assert.StartsWith("[name]", session.GetAnchorInfo(ps[1])!.VisibleText);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("{ MERGEFIELD X")]
    [InlineData("MERGEFIELD X }")]
    [InlineData("IF {} = 1")]
    [InlineData("\"quoted\"")]
    public void IF004_BadCode_Refused(string code)
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var p = BodyParagraphs(session)[0];
        var r = session.InsertField(p, 0, code);
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.InvalidFieldCode, r.Error!.Code);
    }

    [Fact]
    public void IF005_OffsetOutOfRange_Refused()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var r = session.InsertField(BodyParagraphs(session)[0], 9999, "MERGEFIELD X");
        Assert.Equal(EditErrorCode.OffsetOutOfRange, r.Error!.Code);
    }

    [Fact]
    public void IF006_FieldSurvivesSaveAndReopen()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var p = BodyParagraphs(session)[0];
        Assert.True(session.InsertField(p, 0, "MERGEFIELD First").Success);
        using var reopened = new DocxSession(session.Save());
        Assert.StartsWith("«First»", reopened.GetAnchorInfo(BodyParagraphs(reopened)[0])!.VisibleText);
    }

    [Fact]
    public void IF007_StampReferenceFields_MarksMergeFieldsWhole()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var ps = BodyParagraphs(session);
        Assert.True(session.InsertField(ps[0], 0, "MERGEFIELD BoFull").Success);
        Assert.True(session.InsertField(ps[1], 0, @"IF { MERGEFIELD X } = True ""a"" ""b""").Success);
        var bytes = session.Save();

        Assert.DoesNotContain("data-field=\"MERGEFIELD\"", HtmlConversionOps.ConvertToHtml(bytes, new HtmlConversionOptions()));
        var html = XElement.Parse(HtmlConversionOps.ConvertToHtml(bytes, new HtmlConversionOptions { StampReferenceFields = true }));
        var merge = html.Descendants().Single(e => (string?)e.Attribute("data-field") == "MERGEFIELD");
        Assert.Equal("BoFull", (string?)merge.Attribute("data-field-target"));
        Assert.Equal("«BoFull»", merge.Value);
        var iff = html.Descendants().Single(e => (string?)e.Attribute("data-field") == "IF");
        Assert.StartsWith("«IF X = True", iff.Value);
        Assert.Equal(iff.Value, (string?)iff.Attribute("data-field-cached"));
    }

    [Fact]
    public void IF008_StampReferenceFields_CarriesMergeFieldCode()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var ps = BodyParagraphs(session);
        Assert.True(session.InsertField(ps[0], 0, @"MERGEFIELD Amount \# $,0.00").Success);
        Assert.True(session.InsertField(ps[1], 0, @"IF { MERGEFIELD X } = True ""a"" ""b""").Success);
        var html = XElement.Parse(HtmlConversionOps.ConvertToHtml(session.Save(), new HtmlConversionOptions { StampReferenceFields = true }));
        var merge = html.Descendants().Single(e => (string?)e.Attribute("data-field") == "MERGEFIELD");
        Assert.Equal(@"MERGEFIELD Amount \# $,0.00", (string?)merge.Attribute("data-field-code"));
        var iff = html.Descendants().Single(e => (string?)e.Attribute("data-field") == "IF");
        Assert.Matches(@"^IF \{ ?MERGEFIELD X ?\} = True ""a"" ""b""$", (string?)iff.Attribute("data-field-code"));
    }
}
