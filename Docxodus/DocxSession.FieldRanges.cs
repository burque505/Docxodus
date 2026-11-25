// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus;

/// <summary>
/// Complex fields (<c>w:fldChar</c> begin … separate … end) and ranges of text: Word never deletes
/// part of a field's markers. A selection that reaches into a field from outside (a table of
/// contents, say) grows to cover the whole field, and deleting it deletes the whole field.
/// </summary>
public sealed partial class DocxSession
{
    /// <summary>
    /// <paramref name="range"/> grown, as Word grows a selection, to cover every complex field it
    /// covers only part of: a field whose begin is in the range and whose end is not, or the other
    /// way round. A range inside one field's result (some entries of a table of contents) is not
    /// grown. Fields are paired within the range's story (a text box is a story of its own). Null
    /// when an endpoint is not a paragraph of one part, or the range runs backwards.
    /// </summary>
    public DocumentRange? ExpandRangeToFields(DocumentRange range)
    {
        var startPara = FindAnchor(range.StartAnchorId)?.Resolve(_doc!);
        var endPara = FindAnchor(range.EndAnchorId)?.Resolve(_doc!);
        if (startPara is null || endPara is null || startPara.Name != W.p || endPara.Name != W.p) return null;
        var story = StoryRootOf(startPara);
        if (!ReferenceEquals(story, StoryRootOf(endPara))) return null;
        var paragraphs = StoryElements(story, W.p).ToList();
        var index = new Dictionary<XElement, int>(ReferenceEqualityComparer.Instance);
        for (int i = 0; i < paragraphs.Count; i++) index[paragraphs[i]] = i;
        if (!index.TryGetValue(startPara, out var sp) || !index.TryGetValue(endPara, out var ep)) return null;
        var s = (P: sp, O: range.StartOffset);
        var e = (P: ep, O: range.EndOffset);
        if (Compare(s, e) > 0) return null;

        var fields = new List<((int P, int O) Begin, (int P, int O) End)>();
        foreach (var (begin, end) in CompleteFields(story))
        {
            var bp = begin.Ancestors(W.p).FirstOrDefault();
            var fp = end.Ancestors(W.p).FirstOrDefault();
            if (bp is null || fp is null || !index.TryGetValue(bp, out var bi) || !index.TryGetValue(fp, out var fi)) continue;
            fields.Add(((bi, MarkerOffset(bp, begin)), (fi, MarkerOffset(fp, end))));
        }
        bool grown = true;
        while (grown)
        {
            grown = false;
            foreach (var (b, f) in fields)
            {
                bool beginIn = Compare(s, b) <= 0 && Compare(b, e) < 0;
                bool endIn = Compare(s, f) < 0 && Compare(f, e) <= 0;
                if (beginIn == endIn) continue;
                if (Compare(b, s) < 0) { s = b; grown = true; }
                if (Compare(f, e) > 0) { e = f; grown = true; }
            }
        }
        var startAnchor = AnchorForElement(paragraphs[s.P]);
        var endAnchor = AnchorForElement(paragraphs[e.P]);
        if (startAnchor is null || endAnchor is null) return null;
        return new DocumentRange(startAnchor.Value.Id, s.O, endAnchor.Value.Id, e.O);

        static int Compare((int P, int O) a, (int P, int O) b) => a.P != b.P ? a.P.CompareTo(b.P) : a.O.CompareTo(b.O);
    }

    /// <summary>
    /// Take out the field markers left without their partners in the story of
    /// <paramref name="anchorId"/>'s part: an end or a separate with no begin before it, a begin
    /// never ended (with its field code and separate; its result stays as plain text), and field
    /// code outside any field's code. What an edit that cut through a field leaves behind; Word
    /// shows such a document but cannot update or delete the broken field. One undo step, none when
    /// nothing is broken; nothing is done while changes are tracked (deleted markers stay paired
    /// inside their revision).
    /// </summary>
    public EditResult RepairFields(string anchorId)
    {
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        var target = FindAnchor(anchorId);
        var element = target?.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (_trackedChanges == TrackedChangeMode.RenderInline) return new EditResult { Success = true };
        var partRoot = element.AncestorsAndSelf().Last();
        var stories = new[] { partRoot }.Concat(partRoot.Descendants(W.txbxContent)).ToList();
        var broken = stories.SelectMany(UnpairedFieldMarkers).ToList();
        if (broken.Count == 0) return new EditResult { Success = true };
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var touched = new List<XElement>();
            foreach (var marker in broken)
            {
                var run = marker.Parent;
                if (marker.Ancestors(W.p).FirstOrDefault() is { } p && !touched.Contains(p)) touched.Add(p);
                marker.Remove();
                if (run is not null && run.Name == W.r && run.Parent is not null && !run.Elements().Any(x => x.Name != W.rPr))
                    run.Remove();
            }
            InvalidateProjectionCache();
            var modified = touched.Select(p => AnchorForElement(p)).Where(a => a is not null).Select(a => a!.Value)
                .Distinct().ToList();
            return new EditResult { Success = true, Modified = modified };
        }
        catch (Exception ex)
        {
            return FailInternal(ex);
        }
    }

    /// <summary>The complex fields in <paramref name="target"/>'s paragraph (begin and end its own
    /// runs) whose whole, non-empty result lies in one of <paramref name="spans"/>: an edit over them
    /// takes the field. None while changes are tracked (the deleted result stays, marked).</summary>
    private List<FoundField> FieldsWithResultIn(AnchorTarget target, IReadOnlyList<CharSpan> spans)
    {
        var found = new List<FoundField>();
        if (_trackedChanges == TrackedChangeMode.RenderInline || spans.All(s => s.Length == 0)) return found;
        if (target.Resolve(_doc!) is not { } paragraph || paragraph.Name != W.p) return found;
        var map = Internal.RunTextMap.Build(paragraph);
        foreach (var field in ScanFields(paragraph))
        {
            if (!field.Complex || field.Separate is null || field.End is null
                || !ReferenceEquals(field.Begin.Parent, paragraph) || !ReferenceEquals(field.End.Parent, paragraph)) continue;
            var result = map.Segments.Where(s => XNode.DocumentOrderComparer.Compare(s.Run, field.Separate) > 0
                && XNode.DocumentOrderComparer.Compare(s.Run, field.End) < 0).ToList();
            if (result.Count == 0) continue;
            int start = result.Min(s => s.StartOffsetInBlock), end = result.Max(s => s.EndOffsetInBlock);
            if (spans.Any(s => s.Length > 0 && s.Start <= start && end <= s.Start + s.Length)) found.Add(field);
        }
        return found;
    }

    /// <summary>Take out what is left of a field whose result was deleted: its begin, instruction,
    /// separator and end runs (anything else between them, a bookmark say, stays).</summary>
    private static void RemoveFieldMarkers(FoundField field)
    {
        if (field.Begin.Parent is null || field.End?.Parent is null) return;
        var runs = new List<XElement> { field.Begin };
        runs.AddRange(field.Begin.ElementsAfterSelf().TakeWhile(e => !ReferenceEquals(e, field.End))
            .Where(e => e.Name == W.r && e.Elements().All(c => c.Name == W.rPr || c.Name == W.instrText || c.Name == W.fldChar
                || (c.Name == W.t && c.Value.Length == 0))));
        runs.Add(field.End);
        foreach (var run in runs) run.Remove();
    }

    /// <summary>The text box (<c>w:txbxContent</c>) <paramref name="element"/> is in, else its part's root.</summary>
    private static XElement StoryRootOf(XElement element) =>
        element.Ancestors(W.txbxContent).FirstOrDefault() ?? element.AncestorsAndSelf().Last();

    /// <summary>The <paramref name="name"/> elements of the story <paramref name="story"/>, in
    /// document order, leaving out those in a text box inside it (a story of their own).</summary>
    private static IEnumerable<XElement> StoryElements(XElement story, XName name) =>
        story.Descendants(name).Where(e => ReferenceEquals(StoryRootOf(e), story));

    /// <summary>Each complex field of <paramref name="story"/> whose begin has an end: the two
    /// <c>w:fldChar</c> markers. Unpaired markers are skipped.</summary>
    private static IEnumerable<(XElement Begin, XElement End)> CompleteFields(XElement story)
    {
        var open = new Stack<XElement>();
        foreach (var marker in StoryElements(story, W.fldChar))
        {
            switch ((string?)marker.Attribute(W.fldCharType))
            {
                case "begin": open.Push(marker); break;
                case "end":
                    if (open.Count > 0) yield return (open.Pop(), marker);
                    break;
            }
        }
    }

    /// <summary>The markers of <paramref name="story"/> without partners (see <see cref="RepairFields"/>).</summary>
    private static List<XElement> UnpairedFieldMarkers(XElement story)
    {
        var broken = new List<XElement>();
        var open = new Stack<(XElement Begin, List<XElement> Code, XElement? Separate)>();
        foreach (var marker in story.Descendants().Where(e => (e.Name == W.fldChar || e.Name == W.instrText)
            && ReferenceEquals(StoryRootOf(e), story)))
        {
            if (marker.Name == W.instrText)
            {
                if (open.Count > 0 && open.Peek().Separate is null) open.Peek().Code.Add(marker);
                else broken.Add(marker);
                continue;
            }
            switch ((string?)marker.Attribute(W.fldCharType))
            {
                case "begin":
                    open.Push((marker, new List<XElement>(), null));
                    break;
                case "separate":
                    if (open.Count > 0 && open.Peek().Separate is null)
                    {
                        var top = open.Pop();
                        open.Push((top.Begin, top.Code, marker));
                    }
                    else broken.Add(marker);
                    break;
                case "end":
                    if (open.Count > 0) open.Pop();
                    else broken.Add(marker);
                    break;
            }
        }
        foreach (var (begin, code, separate) in open)
        {
            broken.Add(begin);
            broken.AddRange(code);
            if (separate is not null) broken.Add(separate);
        }
        return broken;
    }
}
