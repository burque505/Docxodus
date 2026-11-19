// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus;

// ── Text columns ───────────────────────────────────────────────────────────────────────
//
// Word's Layout ▸ Columns menu and its More Columns dialog. Columns are a section property
// (w:cols), so "Apply to" decides which section gets them: the caret's, every section, a new one
// from the caret on, or a new one around the selected text — the last two put continuous
// section breaks in first, exactly as Word does, so text before and after keeps its columns.

public sealed partial class DocxSession
{
    /// <summary>The scopes <see cref="SetColumns"/> applies to (Word's "Apply to").</summary>
    internal static readonly string[] ColumnScopes = { "section", "forward", "selection", "document" };

    /// <summary>Word allows up to 45 columns (<c>ST_DecimalNumber</c> in practice: its dialog's
    /// limit), and none narrower than half an inch.</summary>
    private const int MaxColumns = 45;
    private const int MinColumnWidthTwips = 720;

    /// <summary>
    /// Set the text columns (<see cref="ColumnsOp"/>) — Word's Layout ▸ Columns.
    /// <paramref name="applyTo"/> is Word's "Apply to":
    /// <list type="bullet">
    /// <item><c>section</c> (the default): the section <paramref name="anchorId"/> is in.</item>
    /// <item><c>document</c>: every section.</item>
    /// <item><c>forward</c> ("This point forward"): a continuous section break goes in at
    /// <paramref name="characterOffset"/>, and the new section after it gets the columns.</item>
    /// <item><c>selection</c> ("Selected text"): continuous section breaks go in around the text
    /// from (<paramref name="anchorId"/>, <paramref name="characterOffset"/>) to
    /// (<paramref name="endAnchorId"/>, <paramref name="endOffset"/>), and only that text gets
    /// the columns.</item>
    /// </list>
    /// </summary>
    /// <remarks>
    /// A break is only added where no section already starts: a selection or caret at the start
    /// of a paragraph puts the break on the paragraph before it, and one that ends at the end of a
    /// paragraph puts it on that paragraph's own mark, so no empty paragraph is left behind.
    /// <c>forward</c> and <c>selection</c> are refused while changes are tracked
    /// (<see cref="EditErrorCode.TrackedOperationUnsupported"/>), as section breaks are; the other
    /// two are tracked as a property change. Either end of a selection may be in a table: the
    /// break then goes before or after the whole table. One undo step.
    /// </remarks>
    public EditResult SetColumns(string anchorId, int characterOffset, ColumnsOp op, string applyTo = "section",
        string? endAnchorId = null, int endOffset = 0)
    {
        ArgumentNullException.ThrowIfNull(op);
        if (!ColumnScopes.Contains(applyTo, StringComparer.Ordinal))
            return EditResult.Fail(EditErrorCode.InvalidPageSetup,
                $"unknown scope '{applyTo}': use {string.Join(", ", ColumnScopes)}", anchorId);
        if (op.Count is { } count && (count < 1 || count > MaxColumns))
            return EditResult.Fail(EditErrorCode.InvalidPageSetup, $"columns must be 1 to {MaxColumns}", anchorId);
        if (op.SpaceTwips is { } space && space < 0)
            return EditResult.Fail(EditErrorCode.InvalidPageSetup, "the space between columns cannot be negative", anchorId);

        try
        {
            return applyTo switch
            {
                "section" => EditGoverningSectPr(anchorId, "SetColumns", sectPr => ApplyColumns(sectPr, op)),
                "document" => SetColumnsEverywhere(anchorId, op),
                _ => SetColumnsInNewSection(anchorId, characterOffset, op, applyTo == "selection",
                    endAnchorId ?? anchorId, endAnchorId is null ? characterOffset : endOffset),
            };
        }
        catch (PageSetupException ex)
        {
            return EditResult.Fail(EditErrorCode.InvalidPageSetup, ex.Message, anchorId);
        }
    }

    /// <summary>Every section of the body (each paragraph's mark that ends one, and the body's
    /// own), in one undo step.</summary>
    private EditResult SetColumnsEverywhere(string anchorId, ColumnsOp op)
    {
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        var body = _doc!.MainDocumentPart!.GetXDocument().Root!.Element(W.body);
        if (body is null) return EditResult.Fail(EditErrorCode.AnchorWrongKind, "the document has no body", anchorId);

        var sections = body.Elements(W.p).Select(p => p.Element(W.pPr)?.Element(W.sectPr))
            .Where(s => s is not null).Cast<XElement>()
            .Concat(body.Elements(W.sectPr)).ToList();
        // Validate every section on a copy first, so a page too narrow anywhere changes nothing.
        bool any = body.Element(W.sectPr) is null && ApplyColumns(new XElement(W.sectPr), op);
        foreach (var sectPr in sections) any |= ApplyColumns(new XElement(sectPr), op);
        var succeeded = new EditResult { Success = true, Modified = new[] { target.Anchor } };
        if (!any) return succeeded;

        using var transaction = BeginTransaction(fullPackage: false);
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            if (body.Element(W.sectPr) is null)
            {
                var final = new XElement(W.sectPr);
                body.Add(final);
                UnidHelper.AssignToSelfAndDescendants(final);
                sections.Add(final);
            }
            foreach (var sectPr in sections)
                if (ApplyColumns(new XElement(sectPr), op))
                    MutateSectionProperties(sectPr, () => ApplyColumns(sectPr, op));
            InvalidateProjectionCache();
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
        transaction.Commit();
        return succeeded;
    }

    /// <summary>"This point forward" and "Selected text": continuous section breaks where the new
    /// section starts (and, for a selection, where it ends), then the columns on that section.</summary>
    private EditResult SetColumnsInNewSection(string anchorId, int offset, ColumnsOp op, bool selection,
        string endAnchorId, int endOffset)
    {
        const string opName = "SetColumns";
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        var (start, startError) = ResolveColumnsEnd(anchorId, offset, opName);
        if (startError is not null) return startError;
        XElement? end = null;
        if (selection)
        {
            var (resolved, endError) = ResolveColumnsEnd(endAnchorId, endOffset, opName);
            if (endError is not null) return endError;
            end = resolved;
            if (start!.IsAfter(end!) || (ReferenceEquals(start, end) && offset > endOffset))
                return EditResult.Fail(EditErrorCode.InvalidPosition, "the selection ends before it starts", anchorId);
        }
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return TrackedStructureUnsupported(opName, anchorId);
        var body = start!.Ancestors(W.body).First();

        // The columns must fit the page of the section the text is in now.
        var governing = Internal.BlockMetadataOps.FindGoverningSectPr(start);
        ApplyColumns(governing is null ? new XElement(W.sectPr) : new XElement(governing), op);

        using var transaction = BeginTransaction(fullPackage: false);
        try
        {
            // Recorded up front: a break that only marks an existing paragraph records nothing of
            // its own, and the transaction must still see the package change (one undo step).
            _history.RecordPreOp(TakeSnapshot());
            var created = new System.Collections.Generic.List<Anchor>();
            // The end first: a split there leaves the start (earlier in the same paragraph, or in
            // an earlier one) where it was.
            if (selection && BreakAfter(end!, endAnchorId, endOffset, body, created) is { } endFailure)
                return endFailure;
            var first = BreakBefore(start, anchorId, offset, body, created, out var startFailure);
            if (startFailure is not null) return startFailure;

            var sectPr = Internal.BlockMetadataOps.FindGoverningSectPr(first!);
            if (sectPr is null)
            {
                sectPr = new XElement(W.sectPr);
                body.Add(sectPr);
                UnidHelper.AssignToSelfAndDescendants(sectPr);
            }
            ApplyColumns(sectPr, op);
            InvalidateProjectionCache();
            transaction.Commit();
            var anchor = FindAnchor(anchorId)?.Anchor;
            return new EditResult
            {
                Success = true,
                Modified = anchor is { } found ? new[] { found } : Array.Empty<Anchor>(),
                Created = created,
            };
        }
        catch (PageSetupException)
        {
            throw;
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    /// <summary>One end of a "This point forward" or "Selected text": a body paragraph, or a
    /// paragraph in a table (the whole table is then the end's block). Text boxes, headers,
    /// footers and notes are refused.</summary>
    private (XElement? Paragraph, EditResult? Error) ResolveColumnsEnd(string anchorId, int offset, string opName)
    {
        var target = FindAnchor(anchorId);
        if (target is null)
            return (null, EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId));
        if (target.Anchor.Scope != "body" || target.Anchor.Kind is not ("p" or "h" or "li"))
            return (null, EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"{opName} requires a paragraph in the document body", anchorId));
        var element = target.Resolve(_doc!);
        if (element is null) return (null, EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId));
        if (element.Ancestors(W.body).FirstOrDefault() is null || element.Ancestors(W.txbxContent).Any())
            return (null, EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"{opName}: a section cannot start or end inside a text box", anchorId));
        var length = ParagraphText(element).Length;
        if (offset < 0 || offset > length)
            return (null, EditResult.Fail(EditErrorCode.OffsetOutOfRange, $"offset {offset} out of [0, {length}]", anchorId));
        return (element, null);
    }

    /// <summary>The body-level block <paramref name="element"/> is in (itself, or its table).</summary>
    private static XElement BodyBlock(XElement element) =>
        element.AncestorsAndSelf().First(e => e.Parent?.Name == W.body);

    /// <summary>
    /// A section must start at (<paramref name="paragraph"/>, <paramref name="offset"/>): end the
    /// one before there, unless one already ends there. Returns the first paragraph of the new
    /// section.
    /// </summary>
    private XElement? BreakBefore(XElement paragraph, string anchorId, int offset, XElement body,
        System.Collections.Generic.List<Anchor> created, out EditResult? failure)
    {
        failure = null;
        var block = BodyBlock(paragraph);
        bool atStart = offset == 0 || !ReferenceEquals(block, paragraph);
        if (atStart)
        {
            var previous = block.ElementsBeforeSelf().LastOrDefault(e => e.Name != W.sectPr);
            // The start of the document is the start of a section already.
            if (previous is null) return paragraph;
            if (previous.Name == W.p)
            {
                if (previous.Element(W.pPr)?.Element(W.sectPr) is null)
                    EndSectionAt(previous, body, "continuous");
                return paragraph;
            }
            if (!ReferenceEquals(block, paragraph))
            {
                // A table right after a table or other block: Word puts the break on an empty
                // paragraph between them.
                var spacer = new XElement(W.p);
                block.AddBeforeSelf(spacer);
                UnidHelper.AssignToSelfAndDescendants(spacer);
                EndSectionAt(spacer, body, "continuous");
                return paragraph;
            }
        }
        var split = SplitParagraph(anchorId, offset);
        if (!split.Success)
        {
            failure = split;
            return null;
        }
        EndSectionAt(paragraph, body, "continuous");
        if (split.Created is { Count: > 0 } made) created.AddRange(made);
        return paragraph.ElementsAfterSelf(W.p).First();
    }

    /// <summary>
    /// A section must end at (<paramref name="paragraph"/>, <paramref name="offset"/>): end it
    /// there, unless one already ends there. A paragraph's end is its mark, unless it is the last
    /// block of the body (the body's own section must keep a paragraph).
    /// </summary>
    private EditResult? BreakAfter(XElement paragraph, string anchorId, int offset, XElement body,
        System.Collections.Generic.List<Anchor> created)
    {
        var block = BodyBlock(paragraph);
        bool lastBlock = !block.ElementsAfterSelf().Any(e => e.Name != W.sectPr);
        if (!ReferenceEquals(block, paragraph))
        {
            // After a table: on an empty paragraph after it, as Word does (and, when the table
            // ends the body, one more for the body's own section to keep).
            var spacer = new XElement(W.p);
            block.AddAfterSelf(spacer);
            UnidHelper.AssignToSelfAndDescendants(spacer);
            EndSectionAt(spacer, body, "continuous");
            if (lastBlock)
            {
                var tail = new XElement(W.p);
                spacer.AddAfterSelf(tail);
                UnidHelper.AssignToSelfAndDescendants(tail);
            }
            return null;
        }
        if (offset == ParagraphText(paragraph).Length && !lastBlock)
        {
            if (paragraph.Element(W.pPr)?.Element(W.sectPr) is null)
                EndSectionAt(paragraph, body, "continuous");
            return null;
        }
        var split = SplitParagraph(anchorId, offset);
        if (!split.Success) return split;
        EndSectionAt(paragraph, body, "continuous");
        if (split.Created is { Count: > 0 } made) created.AddRange(made);
        return null;
    }

    /// <summary>The <see cref="SetColumns"/> mutator: false when the section already has the
    /// columns asked for; throws <see cref="PageSetupException"/> when they do not fit its page.</summary>
    private static bool ApplyColumns(XElement sectPr, ColumnsOp op)
    {
        static int Read(XElement? e, XName attr, int fallback) =>
            WordprocessingMLUtil.TryAttributeToTwips(e?.Attribute(attr)) ?? fallback;
        static string S(int v) => v.ToString(CultureInfo.InvariantCulture);

        var cols = sectPr.Element(W.cols);
        bool unequal = cols?.Elements(W.col).Any() == true;
        int currentCount = int.TryParse((string?)cols?.Attribute(W.num), out var n) ? n
            : unequal ? cols!.Elements(W.col).Count() : 1;
        int count = op.Count ?? currentCount;
        int space = op.SpaceTwips ?? Read(cols, W.space, 720);

        var pgMar = sectPr.Element(W.pgMar);
        int textWidth = Read(sectPr.Element(W.pgSz), W._w, 12240)
            - Read(pgMar, W.left, 1440) - Read(pgMar, W.right, 1440) - Read(pgMar, W.gutter, 0);
        if (count > 1 && (textWidth - space * (count - 1)) / count < MinColumnWidthTwips)
            throw new PageSetupException(
                $"{count} columns {space} twips apart leave less than half an inch for each on a {textWidth}-twip-wide text area");

        bool changed = false;
        if (cols is null)
        {
            if (op.Count is null or 1 && op.SpaceTwips is null && op.Separator is not true) return false;
            cols = new XElement(W.cols);
            WordprocessingMLUtil.InsertSectPrChildInOrder(sectPr, cols);
            changed = true;
        }
        if (op.Count is not null && (unequal || count != currentCount))
        {
            // Equal columns again, as Word's presets make them.
            if (unequal) { cols.Elements(W.col).Remove(); changed = true; }
            changed |= SetAttr(cols, W.equalWidth, null);
            changed |= SetAttr(cols, W.num, count > 1 ? S(count) : null);
        }
        if (op.SpaceTwips is { } wantSpace) changed |= SetAttr(cols, W.space, S(wantSpace));
        if (op.Separator is { } sep) changed |= SetAttr(cols, W.sep, sep ? "1" : null);
        return changed;
    }
}
