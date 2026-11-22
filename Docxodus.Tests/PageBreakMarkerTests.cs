#nullable enable

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary><see cref="WmlToHtmlConverterSettings.MarkPageBreaks"/>: a page break in flow output
/// is marked only when asked, and a line break never is.</summary>
public class PageBreakMarkerTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static byte[] WithBreaks()
    {
        var stream = new MemoryStream();
        stream.Write(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var part = document.MainDocumentPart!;
            var xml = part.GetXDocument();
            var paragraph = xml.Descendants(W + "p").First();
            paragraph.Add(
                new XElement(W + "r", new XElement(W + "br")),
                new XElement(W + "r", new XElement(W + "br", new XAttribute(W + "type", "page"))),
                new XElement(W + "r", new XElement(W + "t", "after")));
            part.PutXDocument();
        }
        return stream.ToArray();
    }

    [Fact]
    public void PB001_PageBreaksAreMarkedOnlyWhenAsked()
    {
        var bytes = WithBreaks();
        var plain = HtmlConversionOps.ConvertToHtml(bytes, new HtmlConversionOptions());
        Assert.DoesNotContain("data-docx-page-break", plain);

        var marked = XDocument.Parse(HtmlConversionOps.ConvertToHtml(bytes, new HtmlConversionOptions { MarkPageBreaks = true }));
        var breaks = marked.Descendants().Where(e => e.Name.LocalName == "br" && e.Attribute("data-docx-break") != null).ToList();
        Assert.Equal(2, breaks.Count);
        Assert.Single(breaks, b => (string?)b.Attribute("data-docx-page-break") == "true");
    }
}
