// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;

namespace Docxodus;

/// <summary>How pasted HTML's own formatting is treated (<see cref="HtmlInsertOptions.Formatting"/>).</summary>
public enum HtmlPasteFormatting
{
    /// <summary>Word's "Merge Formatting": the paragraphs, headings, lists, tables and links, bold,
    /// italic, underline, strikethrough and super/subscript come along; fonts, sizes, colors and
    /// spacing are the document's, taken from the text where the paste lands.</summary>
    Merge,

    /// <summary>Word's "Keep Source Formatting": the source's fonts, sizes, colors and spacing come
    /// along as direct formatting.</summary>
    Keep,
}

/// <summary>Options for <see cref="DocxSession.InsertHtml"/>.</summary>
public sealed record HtmlInsertOptions
{
    public HtmlPasteFormatting Formatting { get; init; } = HtmlPasteFormatting.Merge;

    /// <summary>The source's own style sheet (the text of its <c>&lt;style&gt;</c> elements), applied
    /// as the author style sheet. A <c>&lt;style&gt;</c> inside the HTML is read too.</summary>
    public string? Css { get; init; }
}

/// <summary>Where the caret belongs after an edit: a paragraph anchor and a text offset in it.</summary>
public sealed record TextCaret(string AnchorId, int Offset);

public sealed partial class DocxSession
{
    /// <summary>
    /// Paste HTML (the clipboard's <c>text/html</c>, from Word, a browser or this editor) into the
    /// paragraph <paramref name="anchorId"/> at <paramref name="characterOffset"/>, replacing the
    /// <paramref name="replaceLength"/> characters there, as one undo step. The HTML must be
    /// well-formed XML (XHTML; a browser's <c>XMLSerializer</c> makes it).
    /// <para>
    /// The pasted paragraphs, headings, lists, tables, links and pictures (data: URLs) arrive as
    /// Word's paste puts them: a single plain paragraph goes into the text at the offset; with
    /// more, the paragraph is split there, the first pasted paragraph joins the text before the
    /// offset and the last joins the text after it (a heading, list item or table stays a
    /// paragraph of its own). New styles and list definitions are added to the document; lists keep
    /// their own numbering. Under tracked changes the paste is an insertion.
    /// </para>
    /// <para>The result's <see cref="EditResult.Caret"/> is the end of the pasted content.</para>
    /// </summary>
    public EditResult InsertHtml(string anchorId, int characterOffset, int replaceLength, string html,
        HtmlInsertOptions? options = null)
    {
        if (ValidatePasteTarget(anchorId, characterOffset, replaceLength, "InsertHtml") is { } invalid) return invalid;
        options ??= new HtmlInsertOptions();

        XElement xhtml;
        try { xhtml = XElement.Parse(html, LoadOptions.PreserveWhitespace); }
        catch (System.Xml.XmlException ex)
        {
            return EditResult.Fail(EditErrorCode.MalformedMarkdown, $"the HTML is not well-formed XML: {ex.Message}", anchorId);
        }

        PastedContent pasted;
        try { pasted = ConvertPastedHtml(xhtml, options); }
        catch (Exception ex)
        {
            return EditResult.Fail(EditErrorCode.MalformedMarkdown, $"the HTML could not be converted: {ex.Message}", anchorId);
        }
        if (pasted.Blocks.Count == 0)
            return EditResult.Fail(EditErrorCode.MalformedMarkdown, "the HTML holds nothing to paste", anchorId);
        return PasteBlocks(anchorId, characterOffset, replaceLength, pasted, options.Formatting);
    }

    /// <summary>
    /// Paste plain text (Paste as Plain Text, Keep Text Only) into the paragraph
    /// <paramref name="anchorId"/> at <paramref name="characterOffset"/>, replacing the
    /// <paramref name="replaceLength"/> characters there, as one undo step. Each line is a
    /// paragraph (spliced as <see cref="InsertHtml"/> splices them) and each tab a tab; the text
    /// takes the formatting and paragraph style where it lands. The result's
    /// <see cref="EditResult.Caret"/> is the end of the pasted text.
    /// </summary>
    public EditResult InsertPastedText(string anchorId, int characterOffset, int replaceLength, string text)
    {
        if (ValidatePasteTarget(anchorId, characterOffset, replaceLength, "InsertPastedText") is { } invalid) return invalid;
        text = text.Replace("\r\n", "\n").Replace('\r', '\n');
        if (XmlText.ValidatePayload(text.Replace("\n", "").Replace("\t", ""), anchorId) is { } textError) return EditResult.Fail(textError);
        if (text.Length == 0) return EditResult.Fail(EditErrorCode.MalformedMarkdown, "there is no text to paste", anchorId);
        var blocks = text.Split('\n').Select(line => new XElement(W.p,
            line.Split('\t').SelectMany((piece, i) => new[]
            {
                i > 0 ? new XElement(W.r, new XElement(W.tab)) : null,
                piece.Length > 0 ? new XElement(W.r, new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), piece)) : null,
            }).OfType<XElement>())).ToList();
        var pasted = new PastedContent(blocks, new(), new(), new(), new(StringComparer.Ordinal), new(StringComparer.Ordinal));
        return PasteBlocks(anchorId, characterOffset, replaceLength, pasted, HtmlPasteFormatting.Merge);
    }

    private EditResult? ValidatePasteTarget(string anchorId, int characterOffset, int replaceLength, string op)
    {
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"{op} requires a paragraph/heading/list-item anchor; got kind={target.Anchor.Kind}", anchorId);
        var paragraph = target.Resolve(_doc!);
        if (paragraph is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);
        var length = ParagraphText(paragraph).Length;
        if (characterOffset < 0 || replaceLength < 0 || characterOffset + replaceLength > length)
            return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                $"span [{characterOffset}, {characterOffset + replaceLength}) outside paragraph of length {length}", anchorId);
        return null;
    }

    /// <summary>Delete the selection, split the paragraph, then splice: each step records its own
    /// undo snapshot, and the transaction squashes them into one.</summary>
    private EditResult PasteBlocks(string anchorId, int characterOffset, int replaceLength, PastedContent pasted,
        HtmlPasteFormatting formatting)
    {
        using var transaction = BeginTransaction(fullPackage: false);
        _patchSuppression++;
        EditResult result;
        try
        {
            var id = anchorId;
            if (replaceLength > 0)
            {
                var deleted = ReplaceTextAtSpan(id, characterOffset, replaceLength, "");
                if (!deleted.Success) return deleted;
                if (deleted.Modified.Count > 0) id = deleted.Modified[0].Id;
            }
            result = SplicePasted(id, characterOffset, pasted, formatting);
            if (!result.Success) return result;
        }
        finally { _patchSuppression--; }
        var patchTarget = FindAnchor(result.Modified.Count > 0 ? result.Modified[0].Id : anchorId);
        var final = new EditResult
        {
            Success = true,
            Modified = result.Modified,
            Created = result.Created,
            Caret = result.Caret,
            Patch = patchTarget is null ? null : PatchFor(patchTarget),
        };
        transaction.Commit();
        return final;
    }

    private sealed record PastedContent(
        List<XElement> Blocks,
        List<XElement> NewStyles,
        List<XElement> NewAbstractNums,
        List<XElement> NewNums,
        Dictionary<string, byte[]> Images,
        Dictionary<string, string> Links);

    /// <summary>Run the HTML through <see cref="HtmlToWmlConverter"/> on a small package holding
    /// this document's styles and list definitions (so heading styles are ours and new list ids do
    /// not collide), and collect what it made.</summary>
    private PastedContent ConvertPastedHtml(XElement xhtml, HtmlInsertOptions options)
    {
        var main = _doc!.MainDocumentPart!;
        var ourStyles = main.StyleDefinitionsPart?.GetXDocument().Root;
        var ourNumbering = main.NumberingDefinitionsPart?.GetXDocument().Root;

        WmlDocument start;
        using (var stream = new OpenXmlMemoryStreamDocument(HtmlToWmlConverter.EmptyDocument))
        {
            using (var w = stream.GetWordprocessingDocument())
            {
                var m = w.MainDocumentPart!;
                if (ourStyles is not null)
                    (m.StyleDefinitionsPart ?? m.AddNewPart<StyleDefinitionsPart>()).PutXDocument(new XDocument(WithoutUnids(ourStyles)));
                if (ourNumbering is not null)
                    (m.NumberingDefinitionsPart ?? m.AddNewPart<NumberingDefinitionsPart>()).PutXDocument(new XDocument(WithoutUnids(ourNumbering)));
            }
            start = stream.GetModifiedWmlDocument();
        }

        var settings = HtmlToWmlConverter.GetDefaultSettings();
        var ourSectPr = main.GetXDocument().Root?.Element(W.body)?.Element(W.sectPr);
        if (ourSectPr?.Element(W.pgSz) is not null && ourSectPr.Element(W.pgMar) is not null)
            settings.SectPr = new XElement(W.sectPr, new XElement(ourSectPr.Element(W.pgSz)!), new XElement(ourSectPr.Element(W.pgMar)!));

        var css = string.Join("\n", xhtml.DescendantsAndSelf().Where(e => e.Name.LocalName.Equals("style", StringComparison.OrdinalIgnoreCase))
            .Select(e => e.Value).Append(options.Css ?? ""));
        var converted = HtmlToWmlConverter.ConvertHtmlToWml(PasteUserAgentCss, HtmlToWmlConverter.CleanUpCss(css), "",
            xhtml, settings, start, null!);

        using var ms = new MemoryStream();
        ms.Write(converted.DocumentByteArray, 0, converted.DocumentByteArray.Length);
        using var doc = WordprocessingDocument.Open(ms, false);
        var m2 = doc.MainDocumentPart!;
        var body = m2.GetXDocument().Root!.Element(W.body)!;
        var blocks = body.Elements().Where(e => e.Name == W.p || e.Name == W.tbl).Select(e => new XElement(e)).ToList();

        var styleIds = ourStyles?.Elements(W.style).Select(s => (string?)s.Attribute(W.styleId)).ToHashSet() ?? new HashSet<string?>();
        var newStyles = (m2.StyleDefinitionsPart?.GetXDocument().Root?.Elements(W.style) ?? Enumerable.Empty<XElement>())
            .Where(s => !styleIds.Contains((string?)s.Attribute(W.styleId))).Select(s => new XElement(s)).ToList();

        var abstractIds = ourNumbering?.Elements(W.abstractNum).Select(a => (string?)a.Attribute(W.abstractNumId)).ToHashSet() ?? new HashSet<string?>();
        var numIds = ourNumbering?.Elements(W.num).Select(a => (string?)a.Attribute(W.numId)).ToHashSet() ?? new HashSet<string?>();
        var numbering = m2.NumberingDefinitionsPart?.GetXDocument().Root;
        var usedNums = blocks.SelectMany(b => b.DescendantsAndSelf(W.numId)).Select(n => (string?)n.Attribute(W.val)).ToHashSet();
        var newNums = (numbering?.Elements(W.num) ?? Enumerable.Empty<XElement>())
            .Where(n => !numIds.Contains((string?)n.Attribute(W.numId)) && usedNums.Contains((string?)n.Attribute(W.numId)))
            .Select(n => new XElement(n)).ToList();
        var usedAbstracts = newNums.Select(n => (string?)n.Element(W.abstractNumId)?.Attribute(W.val)).ToHashSet();
        var newAbstracts = (numbering?.Elements(W.abstractNum) ?? Enumerable.Empty<XElement>())
            .Where(a => !abstractIds.Contains((string?)a.Attribute(W.abstractNumId)) && usedAbstracts.Contains((string?)a.Attribute(W.abstractNumId)))
            .Select(a => new XElement(a)).ToList();

        var images = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var embed in blocks.SelectMany(b => b.Descendants(A.blip)).Select(b => (string?)b.Attribute(R.embed)).OfType<string>().Distinct())
        {
            if (m2.GetPartById(embed) is not ImagePart part) continue;
            using var s = part.GetStream();
            using var copy = new MemoryStream();
            s.CopyTo(copy);
            images[embed] = copy.ToArray();
        }
        var links = m2.HyperlinkRelationships.ToDictionary(r => r.Id, r => r.Uri.OriginalString, StringComparer.Ordinal);

        foreach (var block in blocks)
        {
            // A bookmark's name and id could collide with the document's; a pasted one is dropped.
            block.Descendants().Where(e => e.Name == W.bookmarkStart || e.Name == W.bookmarkEnd).Remove();
            foreach (var pPr in block.DescendantsAndSelf(W.pPr)) pPr.Elements(W.sectPr).Remove();
        }
        return new PastedContent(blocks, newStyles, newAbstracts, newNums, images, links);
    }

    private static XElement WithoutUnids(XElement root)
    {
        var copy = new XElement(root);
        foreach (var e in copy.DescendantsAndSelf()) e.Attributes(PtOpenXml.Unid).Remove();
        return copy;
    }

    private static readonly HashSet<XName> MergeKeptRunProperties = new()
    {
        W.rStyle, W.b, W.bCs, W.i, W.iCs, W.u, W.strike, W.dstrike, W.vertAlign, W.caps, W.smallCaps,
    };

    private static readonly HashSet<XName> LandingRunLook = new()
    {
        W.rFonts, W.sz, W.szCs, W.color, W.lang,
    };

    private static readonly HashSet<XName> MergeKeptParagraphProperties = new() { W.pStyle, W.numPr, W.jc };

    /// <summary>Merge Formatting: strip the source's look, keep its structure and emphasis, and give
    /// plain text the look of the run where the paste lands.</summary>
    private void ApplyMergeFormatting(List<XElement> blocks, XElement landing, int offset)
    {
        var landingRun = LandingRunProperties(landing, offset);
        var landingLook = landingRun?.Elements().Where(e => LandingRunLook.Contains(e.Name)).ToList() ?? new List<XElement>();
        var landingStyle = (string?)landing.Element(W.pPr)?.Element(W.pStyle)?.Attribute(W.val);
        bool landingPlain = landing.Element(W.pPr)?.Element(W.numPr) is null
            && WmlToMarkdownConverter.KindFor(landing) is null or "p";

        foreach (var paragraph in blocks.SelectMany(b => b.DescendantsAndSelf(W.p)))
        {
            var pPr = paragraph.Element(W.pPr);
            if (pPr is not null)
            {
                pPr.Elements().Where(e => !MergeKeptParagraphProperties.Contains(e.Name)).Remove();
                if (!pPr.HasElements) pPr.Remove();
            }
            pPr = paragraph.Element(W.pPr);
            bool plain = pPr?.Element(W.pStyle) is null && pPr?.Element(W.numPr) is null;
            // A plain paragraph takes the paragraph style where it lands (body text stays body text).
            if (plain && landingPlain && landingStyle is not null && !paragraph.Ancestors(W.tbl).Any())
            {
                pPr ??= new XElement(W.pPr);
                if (pPr.Parent is null) paragraph.AddFirst(pPr);
                pPr.AddFirst(new XElement(W.pStyle, new XAttribute(W.val, landingStyle)));
            }
            // Every run gets properties to carry the landing look (plain text's runs have none).
            foreach (var run in paragraph.Descendants(W.r).Where(r => r.Element(W.rPr) is null).ToList())
                run.AddFirst(new XElement(W.rPr));
            foreach (var rPr in paragraph.Descendants(W.rPr).ToList())
            {
                rPr.Elements().Where(e => !MergeKeptRunProperties.Contains(e.Name)).Remove();
                // A heading's style makes it bold already; the converter's bold is the browser's h1-h6 default.
                if (pPr?.Element(W.pStyle) is not null && pPr.Element(W.numPr) is null)
                    rPr.Elements().Where(e => e.Name == W.b || e.Name == W.bCs).Remove();
                bool inLink = rPr.Ancestors(W.hyperlink).Any() || rPr.Element(W.rStyle) is not null;
                // The run's own heading or list style decides its look; plain text takes the landing look.
                if (plain && rPr.Parent?.Name == W.r && landingLook.Count > 0)
                    foreach (var look in landingLook)
                        if (!(inLink && look.Name == W.color))
                            WordprocessingMLUtil.InsertRPrChildInOrder(rPr, new XElement(look));
                if (!rPr.HasElements && rPr.Parent?.Name == W.r) rPr.Remove();
            }
        }
        foreach (var table in blocks.SelectMany(b => b.DescendantsAndSelf(W.tbl)))
        {
            table.Element(W.tblPr)?.Elements(W.tblCellSpacing).Remove();
            foreach (var trPr in table.Elements(W.tr).Elements(W.trPr)) trPr.Elements(W.tblCellSpacing).Remove();
        }
    }

    /// <summary>The run properties of the text the paste lands beside: the run before the offset,
    /// else the one after, else the paragraph mark's.</summary>
    private static XElement? LandingRunProperties(XElement paragraph, int offset)
    {
        var map = RunTextMap.Build(paragraph);
        var before = map.Segments.LastOrDefault(s => s.StartOffsetInBlock < offset).Run;
        var after = map.Segments.FirstOrDefault(s => s.StartOffsetInBlock >= offset).Run;
        return (before ?? after)?.Element(W.rPr) ?? paragraph.Element(W.pPr)?.Element(W.rPr);
    }

    private EditResult SplicePasted(string anchorId, int offset, PastedContent pasted, HtmlPasteFormatting formatting)
    {
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        var paragraph = target.Resolve(_doc!);
        if (paragraph is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);
        var length = ParagraphText(paragraph).Length;
        if (HasUnsupportedInlineInsertionBoundary(paragraph, offset))
            return EditResult.Fail(EditErrorCode.UnsupportedInlineBoundary,
                "the paste lands inside a revision or another inline container; paste before or after it", anchorId);

        var blocks = pasted.Blocks;
        if (formatting == HtmlPasteFormatting.Merge) ApplyMergeFormatting(blocks, paragraph, offset);

        static bool IsPlain(XElement block) => block.Name == W.p
            && block.Element(W.pPr)?.Element(W.pStyle) is null && block.Element(W.pPr)?.Element(W.numPr) is null;
        bool inline = blocks.Count == 1 && IsPlain(blocks[0]);

        // In the middle of the text, split first (its own undo snapshot; the transaction squashes it).
        XElement head = paragraph, tail = paragraph;
        bool split = !inline && offset > 0 && offset < length;
        if (split)
        {
            if (SplitWouldCutAField(paragraph, offset))
                return EditResult.Fail(EditErrorCode.UnsupportedInlineBoundary,
                    "the paste lands between a field's begin and end marks; paste before or after the field", anchorId);
            var splitResult = SplitParagraph(anchorId, offset);
            if (!splitResult.Success) return splitResult;
            tail = FindAnchor(splitResult.Created[0].Id)!.Resolve(_doc!)!;
            head = paragraph;
        }

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            MergePastedDefinitions(pasted);
            // Pictures were numbered by the converter from 1; each gets an id the document leaves free.
            var usedDrawingIds = OwnedPartRelationships.StoryParts(_doc!)
                .SelectMany(o => o.Part.GetXDocument().Descendants(WP.docPr))
                .Select(d => uint.TryParse((string?)d.Attribute("id"), out var n) ? n : 0u)
                .ToHashSet();
            void PrepareForDocument(IEnumerable<XElement> content, PastedContent p, OpenXmlPart part) =>
                PreparePastedContent(content, p, part, usedDrawingIds);
            var tracked = _trackedChanges == TrackedChangeMode.RenderInline;
            var stamp = tracked ? new RevisionStamp(_revisionAuthor ?? "docxodus", NextTrackedFormatRevisionDate()) : default;
            var owner = OwnedPartRelationships.FindOwner(_doc!, paragraph)
                ?? throw new InvalidOperationException("paragraph has no owning story part");

            // Inline content of a pasted paragraph, ready to join an existing one.
            List<XElement> InlineOf(XElement p)
            {
                var holder = new XElement(W.p, p.Nodes().Where(n => n is not XElement e || e.Name != W.pPr));
                if (tracked) MarkParagraphContentAndMark(holder, W.ins, stamp.Author, stamp.Date, preserveParagraphMark: true);
                var list = holder.Elements().Where(e => e.Name != W.pPr).ToList();
                foreach (var e in list) e.Remove();   // detached, so each lands in the document itself
                return list;
            }
            static int TextLength(IEnumerable<XElement> inlines) =>
                ParagraphText(new XElement(W.p, inlines.Select(e => new XElement(e)))).Length;

            var created = new List<XElement>();
            TextCaret? caret;

            if (inline)
            {
                var inlines = InlineOf(blocks[0]);
                PrepareForDocument(inlines, pasted, owner.Part);
                InsertInlineSequence(paragraph, offset, inlines);
                caret = new TextCaret(anchorId, offset + TextLength(inlines));
            }
            else
            {
                var middle = new List<XElement>(blocks);
                XElement? anchorBefore = null;   // the inserted blocks go after this one, or before `tail`
                int caretOffset = 0;
                XElement? caretParagraph = null;

                if (offset == 0 && !split)
                {
                    // At the start: everything goes before the paragraph, which stays the tail.
                    head = null!;
                }
                else
                {
                    // At the end, or split: the first plain paragraph joins the head's end.
                    if (IsPlain(middle[0]))
                    {
                        var inlines = InlineOf(middle[0]);
                        PrepareForDocument(inlines, pasted, owner.Part);
                        var headLength = ParagraphText(head).Length;
                        foreach (var e in inlines) head.Add(e);
                        middle.RemoveAt(0);
                        caretParagraph = head;
                        caretOffset = headLength + TextLength(inlines);
                    }
                    anchorBefore = head;
                }

                bool hasTail = head is null || split;
                XElement? tailJoin = null;
                if (hasTail && middle.Count > 0 && middle[^1].Name == W.p && ParagraphText(tail).Length == 0
                    && !tracked && tail.Elements().All(e => e.Name == W.pPr || (e.Name == W.r && !e.Elements().Any(c => c.Name != W.rPr))))
                {
                    // An empty paragraph where the paste lands becomes its last paragraph (no blank
                    // line left behind), keeping the section break it may end.
                    var last = middle[^1];
                    middle.RemoveAt(middle.Count - 1);
                    var sectPr = tail.Element(W.pPr)?.Element(W.sectPr);
                    tail.Elements().Where(e => e.Name != W.pPr).Remove();
                    tail.Element(W.pPr)?.Remove();
                    if (last.Element(W.pPr) is { } lastPPr) tail.AddFirst(new XElement(lastPPr));
                    if (sectPr is not null)
                    {
                        if (tail.Element(W.pPr) is null) tail.AddFirst(new XElement(W.pPr));
                        tail.Element(W.pPr)!.Add(sectPr);
                    }
                    tailJoin = last;
                }
                else if (hasTail && middle.Count > 0 && IsPlain(middle[^1]) && ParagraphText(tail).Length > 0)
                {
                    tailJoin = middle[^1];
                    middle.RemoveAt(middle.Count - 1);
                }

                foreach (var block in middle)
                {
                    PrepareForDocument(new[] { block }, pasted, owner.Part);
                    UnidHelper.AssignToSelfAndDescendants(block);
                    if (anchorBefore is not null) anchorBefore.AddAfterSelf(block);
                    else tail.AddBeforeSelf(block);
                    anchorBefore = block;
                    created.Add(block);
                    if (tracked)
                    {
                        if (block.Name == W.tbl) MarkTableRowsAsTrackedRevision(block, inserted: true, stamp.Author, stamp.Date);
                        else MarkParagraphContentAndMark(block, W.ins, stamp.Author, stamp.Date);
                    }
                }

                if (tailJoin is not null)
                {
                    var inlines = InlineOf(tailJoin);
                    PrepareForDocument(inlines, pasted, owner.Part);
                    InsertInlineSequence(tail, 0, inlines);
                    caretParagraph = tail;
                    caretOffset = TextLength(inlines);
                }
                else if (middle.Count > 0)
                {
                    var last = middle[^1];
                    if (last.Name == W.p) { caretParagraph = last; caretOffset = ParagraphText(last).Length; }
                    else { caretParagraph = hasTail ? tail : null; caretOffset = 0; }
                }

                created.AddRange(SeparateTables(created, tracked, stamp));
                caretParagraph ??= created.LastOrDefault(e => e.Name == W.p) ?? (hasTail ? tail : head);
                UnidHelper.AssignToSelfAndDescendants(caretParagraph!);
                InvalidateProjectionCache();
                _ = AnchorIndex();
                caret = AnchorForElement(caretParagraph!) is { } a ? new TextCaret(a.Id, caretOffset) : null;
            }

            UnidHelper.AssignToSelfAndDescendants(paragraph);
            if (split) UnidHelper.AssignToSelfAndDescendants(tail);
            InvalidateProjectionCache();
            if (inline) caret = AnchorForElement(paragraph) is { } p ? caret! with { AnchorId = p.Id } : caret;

            var modified = new List<Anchor>();
            if (AnchorForElement(paragraph) is { } pa) modified.Add(pa);
            // The split's second half is a new paragraph too.
            var createdElements = created.SelectMany(b => b.DescendantsAndSelf().Where(e => e.Name == W.p || e.Name == W.tc));
            if (split) createdElements = createdElements.Append(tail);
            return new EditResult
            {
                Success = true,
                Modified = modified,
                Created = ResolveAnchorsForElements(createdElements),
                Caret = caret,
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    private void InsertInlineSequence(XElement paragraph, int offset, List<XElement> inlines)
    {
        if (inlines.Count == 0) return;
        InsertInlineElementAtOffset(paragraph, offset, inlines[0]);
        for (int i = 1; i < inlines.Count; i++) inlines[i - 1].AddAfterSelf(inlines[i]);
        foreach (var e in inlines) UnidHelper.AssignToSelfAndDescendants(e);
    }

    /// <summary>Tables must not touch one another or end the story: an empty paragraph goes
    /// between (as <see cref="InsertTable"/> does). Returns the paragraphs added.</summary>
    private List<XElement> SeparateTables(List<XElement> inserted, bool tracked, RevisionStamp stamp)
    {
        var added = new List<XElement>();
        foreach (var table in inserted.Where(b => b.Name == W.tbl).ToList())
        {
            foreach (bool after in new[] { false, true })
            {
                var neighbour = RevisionOps.SiblingBlock(table, forward: after);
                if (neighbour?.Name != W.tbl && !(after && (neighbour is null || neighbour.Name == W.sectPr)))
                    continue;
                var separator = new XElement(W.p);
                UnidHelper.AssignToSelfAndDescendants(separator);
                if (after) table.AddAfterSelf(separator);
                else table.AddBeforeSelf(separator);
                if (tracked) MarkParagraphMark(separator, W.ins, stamp.Author, stamp.Date);
                added.Add(separator);
            }
        }
        return added;
    }

    /// <summary>Add the styles and list definitions the converter made that this document lacks.</summary>
    private void MergePastedDefinitions(PastedContent pasted)
    {
        var main = _doc!.MainDocumentPart!;
        if (pasted.NewStyles.Count > 0 && main.StyleDefinitionsPart is { } stylesPart)
        {
            var root = stylesPart.GetXDocument().Root!;
            foreach (var style in pasted.NewStyles) root.Add(new XElement(style));
            stylesPart.PutXDocument();
        }
        if (pasted.NewNums.Count > 0)
        {
            var part = main.NumberingDefinitionsPart;
            if (part is null)
            {
                part = main.AddNewPart<NumberingDefinitionsPart>();
                part.PutXDocument(new XDocument(new XElement(W.numbering, new XAttribute(XNamespace.Xmlns + "w", W.w.NamespaceName))));
            }
            var root = part.GetXDocument().Root!;
            foreach (var a in pasted.NewAbstractNums) WordprocessingMLUtil.InsertNumberingChildInOrder(root, new XElement(a));
            foreach (var n in pasted.NewNums) WordprocessingMLUtil.InsertNumberingChildInOrder(root, new XElement(n));
            part.PutXDocument();
        }
    }

    /// <summary>Give pasted pictures and links this document's parts, relationships and free
    /// drawing ids.</summary>
    private void PreparePastedContent(IEnumerable<XElement> content, PastedContent pasted, OpenXmlPart owner,
        HashSet<uint> usedDrawingIds)
    {
        foreach (var root in content)
        {
            foreach (var docPr in root.DescendantsAndSelf(WP.docPr))
            {
                uint next = 1;
                while (usedDrawingIds.Contains(next)) next++;
                docPr.SetAttributeValue("id", next);
                usedDrawingIds.Add(next);
            }
            foreach (var blip in root.DescendantsAndSelf(A.blip).ToList())
            {
                var embed = (string?)blip.Attribute(R.embed);
                if (embed is null || !pasted.Images.TryGetValue(embed, out var bytes)) continue;
                var data = ValidateImageBytes(bytes, null);
                if (data.Error is not null || data.ContentType is null)
                {
                    blip.Ancestors(W.r).FirstOrDefault()?.Remove();
                    continue;
                }
                var relationship = OwnedPartRelationships.FindOrAddImagePart(_doc!, owner, bytes, data.ContentType, data.Format);
                blip.SetAttributeValue(R.embed, relationship.RelationshipId);
            }
            foreach (var link in root.DescendantsAndSelf(W.hyperlink))
            {
                var id = (string?)link.Attribute(R.id);
                link.Attribute(R.id)?.Remove();
                if (id is not null && pasted.Links.TryGetValue(id, out var url))
                {
                    if (url.StartsWith("#", StringComparison.Ordinal)) continue;   // a jump to the source's own bookmark
                    var relationship = OwnedPartRelationships.FindOrAddExternalHyperlink(owner, PackageOpenRepair.TargetUri(url));
                    link.SetAttributeValue(R.id, relationship.Id);
                }
            }
        }
    }

    /// <summary>The browser defaults the converter needs (it applies no style sheet of its own),
    /// with tables collapsed as pasted tables are drawn in Word.</summary>
    private const string PasteUserAgentCss = @"html, address, blockquote, body, dd, div, dl, dt, fieldset, form, frame, frameset,
h1, h2, h3, h4, h5, h6, noframes, ol, p, ul, center, dir, hr, menu, pre { display: block; unicode-bidi: embed }
li { display: list-item }
head { display: none }
table { display: table }
tr { display: table-row }
thead { display: table-header-group }
tbody { display: table-row-group }
tfoot { display: table-footer-group }
col { display: table-column }
colgroup { display: table-column-group }
td, th { display: table-cell }
caption { display: table-caption }
th { font-weight: bolder; text-align: center }
caption { text-align: center }
body { margin: auto; }
h1 { font-size: 2em; margin: auto; }
h2 { font-size: 1.5em; margin: auto; }
h3 { font-size: 1.17em; margin: auto; }
h4, p, blockquote, ul, fieldset, form, ol, dl, dir, menu { margin: auto }
a { color: blue; }
h5 { font-size: .83em; margin: auto }
h6 { font-size: .75em; margin: auto }
h1, h2, h3, h4, h5, h6, b, strong { font-weight: bolder }
blockquote { margin-left: 40px; margin-right: 40px }
i, cite, em, var, address { font-style: italic }
pre, tt, code, kbd, samp { font-family: monospace }
pre { white-space: pre }
button, textarea, input, select { display: inline-block }
big { font-size: 1.17em }
small, sub, sup { font-size: .83em }
sub { vertical-align: sub }
sup { vertical-align: super }
table { border-spacing: 0; border-collapse: collapse }
thead, tbody, tfoot { vertical-align: middle }
td, th, tr { vertical-align: inherit }
s, strike, del { text-decoration: line-through }
hr { border: 1px inset }
ol, ul, dir, menu, dd { margin-left: 40px }
ol { list-style-type: decimal }
ol ul, ul ol, ul ul, ol ol { margin-top: 0; margin-bottom: 0 }
u, ins { text-decoration: underline }
br:before { content: ""\A""; white-space: pre-line }
center { text-align: center }
:link, :visited { text-decoration: underline }
BDO[DIR=""ltr""] { direction: ltr; unicode-bidi: bidi-override }
BDO[DIR=""rtl""] { direction: rtl; unicode-bidi: bidi-override }
*[DIR=""ltr""] { direction: ltr; unicode-bidi: embed }
*[DIR=""rtl""] { direction: rtl; unicode-bidi: embed }
";
}
