// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Docxodus.Internal;

namespace Docxodus;

/// <summary>Where <see cref="DocxSession.InsertCaption"/> puts a caption.</summary>
public enum CaptionPosition
{
    /// <summary>Before the item (Word's default for a table).</summary>
    Above,

    /// <summary>After the item (Word's default for anything else, a picture among them).</summary>
    Below,
}

/// <summary>Options for <see cref="DocxSession.InsertCaption"/>: Word's Insert Caption dialog.</summary>
public sealed record CaptionOptions
{
    /// <summary>The caption label ("Figure", "Table", "Equation", or one of the user's own). The
    /// caption's <c>SEQ</c> field counts captions by it, with spaces written as underscores (as
    /// Word does), and a table of figures lists them by it.</summary>
    public string Label { get; init; } = "Figure";

    /// <summary>Above or below the item; null takes Word's default: above a table, below anything
    /// else.</summary>
    public CaptionPosition? Position { get; init; }

    /// <summary>Leave the label out of the caption: only its number (Word's "Exclude label from
    /// caption"). The <c>SEQ</c> field still counts by the label.</summary>
    public bool ExcludeLabel { get; init; }

    /// <summary>The number's format, a <c>\*</c> switch of the <c>SEQ</c> field: <c>ARABIC</c> (1, 2,
    /// 3), <c>alphabetic</c> (a, b, c), <c>ALPHABETIC</c>, <c>roman</c> (i, ii, iii) or
    /// <c>ROMAN</c>.</summary>
    public string NumberFormat { get; init; } = "ARABIC";

    /// <summary>What follows the number, as typed (": The results", say); null or empty for
    /// none.</summary>
    public string? Text { get; init; }
}

/// <summary>Which part of a caption a cross-reference shows (see
/// <see cref="DocxSession.AddReferenceBookmark(string, ReferenceTargetPart)"/>).</summary>
public enum ReferenceTargetPart
{
    /// <summary>The whole paragraph: a heading's or numbered item's text, an entire caption.</summary>
    Paragraph,

    /// <summary>A caption's label and number ("Figure 1"): from the paragraph's start to the end of
    /// its <c>SEQ</c> field.</summary>
    LabelAndNumber,

    /// <summary>A caption's text after its number and the punctuation and spaces that follow it.</summary>
    CaptionText,
}

/// <summary>One caption in the body: a paragraph holding a <c>SEQ</c> field (the first one in it is
/// the caption's number).</summary>
/// <param name="AnchorId">The paragraph's anchor.</param>
/// <param name="Label">The <c>SEQ</c> field's identifier ("Figure", "My_Label").</param>
/// <param name="Number">The number the field shows.</param>
/// <param name="Text">The paragraph's text ("Figure 1: The results").</param>
public sealed record CaptionInfo(string AnchorId, string Label, string Number, string Text);

public sealed partial class DocxSession
{
    private static readonly string[] CaptionNumberFormats = { "ARABIC", "alphabetic", "ALPHABETIC", "roman", "ROMAN" };

    /// <summary>
    /// Word's Insert Caption: a new paragraph in the built-in <c>Caption</c> style (the document's
    /// own, found by its name, else Word's) holding the label, a <c>SEQ</c> field numbering it and
    /// <see cref="CaptionOptions.Text"/>, above or below <paramref name="anchorId"/> (a paragraph,
    /// say one holding a picture, or a table). Every <c>SEQ</c> field in the body is numbered again,
    /// so the captions after the new one move up by one. One undo step. Body blocks only; refused
    /// while changes are tracked, as the reference tables are.
    /// </summary>
    /// <returns>The caption paragraph's anchor in <see cref="EditResult.Created"/>.</returns>
    public EditResult InsertCaption(string anchorId, CaptionOptions? options = null)
    {
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        var o = options ?? new CaptionOptions();
        var label = (o.Label ?? "").Trim();
        if (label.Length == 0 || label.Length > 40 || label.IndexOfAny(new[] { '"', '\\' }) >= 0)
            return EditResult.Fail(EditErrorCode.InvalidReferenceField,
                "Label must be 1 to 40 characters, without quotation marks or backslashes", anchorId);
        // (the case of a letter or roman format is its meaning; arabic has none)
        var format = string.Equals(o.NumberFormat, "arabic", StringComparison.OrdinalIgnoreCase) ? "ARABIC"
            : CaptionNumberFormats.FirstOrDefault(f => f == o.NumberFormat);
        if (format is null)
            return EditResult.Fail(EditErrorCode.InvalidReferenceField,
                $"NumberFormat must be one of {string.Join(", ", CaptionNumberFormats)}; got {o.NumberFormat}", anchorId);
        if ((XmlText.ValidatePayload(label, anchorId) ?? XmlText.ValidatePayload(o.Text ?? "", anchorId)) is { } textError)
            return EditResult.Fail(textError);

        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Scope != "body")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"InsertCaption requires a body anchor; got scope '{target.Anchor.Scope}'", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null || (element.Name != W.p && element.Name != W.tbl))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                "InsertCaption requires a paragraph or table anchor", anchorId);
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return TrackedStructureUnsupported("InsertCaption", anchorId);

        var position = o.Position ?? (element.Name == W.tbl ? CaptionPosition.Above : CaptionPosition.Below);
        var identifier = label.Replace(' ', '_');
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var style = CaptionStyleId();
            var pPr = new XElement(W.pPr, new XElement(W.pStyle, new XAttribute(W.val, style)),
                position == CaptionPosition.Above ? new XElement(W.keepNext) : null);
            var caption = new XElement(W.p, pPr,
                o.ExcludeLabel ? null : new XElement(W.r, TextContent(label + " ", W.t)),
                new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin"))),
                new XElement(W.r, new XElement(W.instrText, new XAttribute(XNamespace.Xml + "space", "preserve"),
                    $" SEQ {identifier} \\* {format} ")),
                new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "separate"))),
                new XElement(W.r, new XElement(W.rPr, new XElement(W.noProof)), TextContent("1", W.t)),
                new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "end"))),
                string.IsNullOrEmpty(o.Text) ? null : new XElement(W.r, TextContent(o.Text, W.t)));
            UnidHelper.AssignToSelfAndDescendants(caption);

            if (position == CaptionPosition.Above) element.AddBeforeSelf(caption);
            else
            {
                element.AddAfterSelf(caption);
                // (the section the item ends stays after its caption)
                if (element.Name == W.p && element.Element(W.pPr)?.Element(W.sectPr) is { } section)
                {
                    section.Remove();
                    SetPPrChildInOrder(pPr, section);
                }
            }

            var body = _doc!.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
            var update = new FieldUpdate(new Dictionary<string, string>(), () => { });
            UpdateSequences(body, ScanFields(body), null, update);
            InvalidateProjectionCache();

            var anchors = BodyAnchors();
            var created = AnchorIn(anchors, caption);
            var modified = update.Modified.Where(p => !ReferenceEquals(p, caption)).Distinct()
                .Select(p => AnchorIn(anchors, p)).OfType<Anchor>().ToList();
            if (element.Name == W.p && AnchorIn(anchors, element) is { } item) modified.Add(item);
            return new EditResult
            {
                Success = true,
                Created = created is { } made ? new[] { made } : Array.Empty<Anchor>(),
                Modified = modified.ToArray(),
                Patch = FindAnchor(anchorId) is { } touched ? PatchFor(touched) : null,
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    /// <summary>The body's captions in document order: each paragraph holding a <c>SEQ</c> field with
    /// an identifier, outside tables of contents and figures. What a Cross-reference dialog lists for a
    /// caption label, and which labels a document already uses.</summary>
    public IReadOnlyList<CaptionInfo> ListCaptions()
    {
        if (_doc?.MainDocumentPart?.GetXDocument().Root?.Element(W.body) is not { } body) return Array.Empty<CaptionInfo>();
        var fields = ScanFields(body);
        var inTables = new HashSet<XElement>(fields.Where(f => f.Kind == "TOC").SelectMany(f => f.Paragraphs()));
        var anchors = BodyAnchors();
        var captions = new List<CaptionInfo>();
        var seen = new HashSet<XElement>();
        foreach (var field in fields)
        {
            if (CaptionIdentifier(field) is not { } identifier) continue;
            var paragraph = field.Paragraph;
            if (inTables.Contains(paragraph) || !seen.Add(paragraph) || AnchorIn(anchors, paragraph) is not { } anchor) continue;
            captions.Add(new CaptionInfo(anchor.Id, identifier, field.ResultText(),
                System.Text.RegularExpressions.Regex.Replace(ParagraphText(paragraph), @"\s+", " ").Trim()));
        }
        return captions;
    }

    /// <summary>
    /// The hidden <c>_Ref</c> bookmark a cross-reference to <paramref name="part"/> of a paragraph
    /// targets: the whole paragraph (as <see cref="AddReferenceBookmark(string)"/>), or a caption's
    /// label and number, or its text alone, as Word's Cross-reference dialog makes for "Only label
    /// and number" and "Only caption text". One there already is reused; otherwise one undo step.
    /// </summary>
    public EditResult AddReferenceBookmark(string anchorId, ReferenceTargetPart part)
    {
        if (part == ReferenceTargetPart.Paragraph) return AddReferenceBookmark(anchorId);
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        var paragraph = target.Resolve(_doc!);
        if (paragraph is null || paragraph.Name != W.p)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                "AddReferenceBookmark requires a paragraph/heading/list-item anchor", anchorId);
        var field = ScanFields(paragraph).FirstOrDefault(f => CaptionIdentifier(f) is not null);
        if (field?.End is null || !ReferenceEquals(field.End.Parent, paragraph) || !ReferenceEquals(field.Begin.Parent, paragraph))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                "the paragraph is not a caption: it holds no SEQ field directly", anchorId);

        var map = RunTextMap.Build(paragraph);
        int textStart = CaptionTextStart(map, field, out int afterNumber);
        int start = part == ReferenceTargetPart.LabelAndNumber ? 0 : textStart;
        int end = part == ReferenceTargetPart.LabelAndNumber ? afterNumber : map.FlatText.Length;
        if (part == ReferenceTargetPart.CaptionText && start >= end)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "the caption has no text after its number", anchorId);

        if (ListBookmarks().FirstOrDefault(b => b.Name.StartsWith("_Ref", StringComparison.Ordinal) && b.IsPaired && b.IsValid
                && b.Range is { } r && r.StartAnchorId == target.Anchor.Id && r.EndAnchorId == target.Anchor.Id
                && r.StartOffset == start && r.EndOffset == end) is { } existing)
            return new EditResult { Success = true, BookmarkName = existing.Name, Modified = new[] { target.Anchor } };

        var name = NewHiddenBookmarkName("_Ref");
        var common = ValidateBookmarkMutation(name, requireValidName: false);
        if (common is not null) return common;
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            {
                // (ends at the field's own end, so the bookmark holds the whole field and never cuts it)
                var id = NextGlobalBookmarkId();
                var bookmarkStart = new XElement(W.bookmarkStart, new XAttribute(W.id, id), new XAttribute(W.name, name));
                var bookmarkEnd = new XElement(W.bookmarkEnd, new XAttribute(W.id, id));
                if (part == ReferenceTargetPart.LabelAndNumber)
                {
                    if (paragraph.Element(W.pPr) is { } pPr) pPr.AddAfterSelf(bookmarkStart);
                    else paragraph.AddFirst(bookmarkStart);
                    field.End.AddAfterSelf(bookmarkEnd);
                }
                else
                {
                    // (after the punctuation and spaces that follow the number: a run holding them is split)
                    XElement after = field.End;
                    int skip = start - afterNumber;
                    while (skip > 0 && after.ElementsAfterSelf().FirstOrDefault() is { } next)
                    {
                        after = next;
                        if (next.Name != W.r) continue;
                        var text = string.Concat(next.Elements().Select(TextOf));
                        if (text.Length <= skip) { skip -= text.Length; continue; }
                        if (next.Elements(W.t).Count() != 1 || next.Elements().Any(e => e.Name != W.t && e.Name != W.rPr)) break;
                        var rest = new XElement(next);
                        next.Element(W.t)!.ReplaceWith(TextContent(text[..skip], W.t));
                        rest.Element(W.t)!.ReplaceWith(TextContent(text[skip..], W.t));
                        UnidHelper.AssignToSelfAndDescendants(rest);
                        next.AddAfterSelf(rest);
                        skip = 0;
                    }
                    after.AddAfterSelf(bookmarkStart);
                    paragraph.Add(bookmarkEnd);
                }
            }
            InvalidateProjectionCache();
            return new EditResult { Success = true, BookmarkName = name, Modified = new[] { target.Anchor } };
        }
        catch (Exception ex)
        {
            return FailInternal(ex);
        }
    }

    /// <summary>Where a caption's own text starts in its paragraph's text: after its number field
    /// (<paramref name="afterNumber"/>) and the punctuation and spaces that follow it.</summary>
    private static int CaptionTextStart(RunTextMap.Map map, FoundField number, out int afterNumber)
    {
        var end = number.End ?? number.Begin;
        afterNumber = map.Segments.Where(s => XNode.DocumentOrderComparer.Compare(s.Run, end) < 0)
            .Select(s => s.EndOffsetInBlock).DefaultIfEmpty(0).Max();
        int start = afterNumber;
        while (start < map.FlatText.Length && (char.IsWhiteSpace(map.FlatText[start]) || ":.-–—".Contains(map.FlatText[start])))
            start++;
        return start;
    }

    /// <summary>A caption's text without its label and number (a table of figures' <c>\a</c>).</summary>
    private static string CaptionTextAfterNumber(XElement paragraph, FoundField number)
    {
        var map = RunTextMap.Build(paragraph);
        return map.FlatText[CaptionTextStart(map, number, out _)..];
    }

    /// <summary>The identifier of a <c>SEQ</c> field that numbers something (not one that only repeats
    /// a bookmark's number, or hides itself), else null.</summary>
    private static string? CaptionIdentifier(FoundField field) =>
        field.Kind == "SEQ" && field.Tokens.Length >= 2 && !field.Tokens[1].StartsWith('\\')
            && !(field.Tokens.Length > 2 && !field.Tokens[2].StartsWith('\\'))
            && !field.Tokens.Skip(2).Any(t => string.Equals(t, "\\h", StringComparison.OrdinalIgnoreCase))
            ? field.Tokens[1] : null;

    /// <summary>The document's caption style (found by its built-in name, so a localized id counts),
    /// else Word's, added.</summary>
    private string CaptionStyleId()
    {
        if (ParagraphStyles().FirstOrDefault(s =>
                string.Equals((string?)s.Element(W.name)?.Attribute(W.val), "caption", StringComparison.OrdinalIgnoreCase))
            ?.Attribute(W.styleId)?.Value is { } own)
            return own;
        StyleFactory.EnsureParagraphStyle(_doc!, "Caption");
        return "Caption";
    }
}
