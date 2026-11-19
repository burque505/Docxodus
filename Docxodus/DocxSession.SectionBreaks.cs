// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus;

// ── Section breaks and break characters ─────────────────────────────────────────────────
//
// Word's Layout ▸ Breaks menu. A section break is a w:sectPr on a paragraph mark: it ENDS the
// section the paragraph is in and holds that section's properties. The section after it is
// governed by the next break (eventually the body's trailing w:sectPr), and that later w:sectPr's
// w:type says how the new section starts (ECMA-376 §17.6.22) — so inserting a break copies the
// governing properties onto the new mark and writes the requested start on the governing one.
// A page or column break is a w:br in a run; Word (2013 on) follows it with a paragraph mark, so
// the text after the caret starts the next page or column as its own paragraph.

public sealed partial class DocxSession
{
    /// <summary>The section starts Word offers (<c>ST_SectionMark</c>): the value of
    /// <c>w:type</c>; <c>nextPage</c> is the default Word writes no element for.</summary>
    internal static readonly string[] SectionStarts = { "nextPage", "continuous", "evenPage", "oddPage", "nextColumn" };

    /// <summary>
    /// Insert a section break at <paramref name="characterOffset"/> in a body paragraph — Word's
    /// Layout ▸ Breaks ▸ Section Breaks. The paragraph is split there; the first half's mark ends
    /// the section with a copy of the properties that governed it (size, margins, columns, header
    /// and footer references, page numbering), and the section after the break starts as
    /// <paramref name="sectionStart"/> says: <c>nextPage</c>, <c>continuous</c>, <c>evenPage</c>,
    /// <c>oddPage</c> or <c>nextColumn</c>.
    /// </summary>
    /// <remarks>
    /// A page-number restart (<c>w:pgNumType/@w:start</c>) stays with the earlier section, where it
    /// was, so the new section numbers on. Refused inside a table, a text box or a header, footer
    /// or note (<see cref="EditErrorCode.AnchorWrongKind"/>), for an unknown start
    /// (<see cref="EditErrorCode.InvalidPageSetup"/>), and while changes are tracked
    /// (<see cref="EditErrorCode.TrackedOperationUnsupported"/>). One undo step.
    /// <see cref="EditResult.Created"/> holds the paragraph that starts the new section.
    /// </remarks>
    public EditResult InsertSectionBreak(string anchorId, int characterOffset, string sectionStart)
    {
        if (!SectionStarts.Contains(sectionStart, StringComparer.Ordinal))
            return EditResult.Fail(EditErrorCode.InvalidPageSetup,
                $"unknown section start '{sectionStart}': use {string.Join(", ", SectionStarts)}", anchorId);
        return SplitWithBreak(anchorId, characterOffset, "InsertSectionBreak",
            (paragraph, body) => EndSectionAt(paragraph, body, sectionStart));
    }

    /// <summary>
    /// Make <paramref name="paragraph"/>'s mark end its section: it gets a copy of the properties
    /// that governed it, and the section after it (governed by the old properties) starts as
    /// <paramref name="sectionStart"/> says. A page-number restart stays with the earlier section.
    /// </summary>
    private static void EndSectionAt(XElement paragraph, XElement body, string sectionStart)
    {
        var governing = Internal.BlockMetadataOps.FindGoverningSectPr(paragraph);
        if (governing is null)
        {
            governing = new XElement(W.sectPr);
            body.Add(governing);
            UnidHelper.AssignToSelfAndDescendants(governing);
        }

        var ending = new XElement(governing);
        ending.Elements(W.sectPrChange).Remove();
        foreach (var el in ending.DescendantsAndSelf())
        {
            el.Attributes().Where(a => a.Name.Namespace == W.w && a.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal)).Remove();
            el.Attributes(PtOpenXml.Unid).Remove();
        }
        UnidHelper.AssignToSelfAndDescendants(ending);
        var pPr = paragraph.Element(W.pPr);
        if (pPr is null)
        {
            pPr = new XElement(W.pPr);
            paragraph.AddFirst(pPr);
        }
        SetPPrChildInOrder(pPr, ending);

        // The new section numbers on from the one before it.
        if (governing.Element(W.pgNumType) is { } pgNumType && pgNumType.Attribute(W.start) is { } start)
        {
            start.Remove();
            if (!pgNumType.Attributes().Any(a => a.Name.Namespace == W.w)) pgNumType.Remove();
        }
        WriteSectionStart(governing, sectionStart);
    }

    /// <summary>
    /// Insert a page break or a column break at <paramref name="characterOffset"/> in a body
    /// paragraph — Word's Ctrl+Enter and Ctrl+Shift+Enter. As Word does, the break ends the first
    /// half of the paragraph (a run holding <c>w:br w:type="page"</c> or <c>"column"</c>) and the
    /// text after the caret goes on as a paragraph of its own. <paramref name="kind"/> is
    /// <c>page</c> or <c>column</c>; refusals as <see cref="InsertSectionBreak"/>. One undo step.
    /// <see cref="EditResult.Created"/> holds the paragraph after the break.
    /// </summary>
    public EditResult InsertBreak(string anchorId, int characterOffset, string kind)
    {
        if (kind is not ("page" or "column"))
            return EditResult.Fail(EditErrorCode.InvalidPosition,
                $"unknown break '{kind}': use page or column", anchorId);
        return SplitWithBreak(anchorId, characterOffset, "InsertBreak", (paragraph, _) =>
        {
            // The break takes the formatting of the text before it, as typed text would.
            var before = paragraph.Elements(W.r).LastOrDefault()?.Element(W.rPr);
            var run = new XElement(W.r,
                before is null ? null : new XElement(before),
                new XElement(W.br, new XAttribute(W.type, kind)));
            run.Element(W.rPr)?.Elements().Where(e => Internal.RevisionOps.RevWrapperNames.Contains(e.Name) || e.Name == W.rPrChange).Remove();
            foreach (var el in run.DescendantsAndSelf()) el.Attributes(PtOpenXml.Unid).Remove();
            UnidHelper.AssignToSelfAndDescendants(run);
            var last = paragraph.Elements().LastOrDefault(e => e.Name != W.pPr);
            if (last is null) paragraph.Add(run);
            else last.AddAfterSelf(run);
        });
    }

    /// <summary>The shared body of the two break inserts: validate, then split the paragraph and
    /// let <paramref name="mark"/> finish the first half — all one transaction, one undo step.</summary>
    private EditResult SplitWithBreak(string anchorId, int characterOffset, string opName,
        Action<XElement, XElement> mark)
    {
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Scope != "body" || target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"{opName} requires a paragraph in the document body", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);
        var body = element.Ancestors(W.body).FirstOrDefault();
        if (body is null || element.Ancestors(W.tbl).Any() || element.Ancestors(W.txbxContent).Any())
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"{opName}: a break cannot go inside a table or a text box", anchorId);
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return TrackedStructureUnsupported(opName, anchorId);

        using var transaction = BeginTransaction(fullPackage: false);
        var split = SplitParagraph(anchorId, characterOffset);
        if (!split.Success) return split;
        try
        {
            mark(element, body);
            InvalidateProjectionCache();
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
        transaction.Commit();
        return new EditResult
        {
            Success = true,
            Modified = new[] { target.Anchor },
            Created = split.Created,
        };
    }

    /// <summary>Write a section's start (<c>w:type</c>); <c>nextPage</c>, the default, removes
    /// the element as Word does. True when the section changed.</summary>
    private static bool WriteSectionStart(XElement sectPr, string sectionStart)
    {
        var existing = sectPr.Element(W.type);
        if (sectionStart == "nextPage")
        {
            if (existing is null) return false;
            existing.Remove();
            return true;
        }
        if (existing is not null)
        {
            if ((string?)existing.Attribute(W.val) == sectionStart) return false;
            existing.SetAttributeValue(W.val, sectionStart);
            return true;
        }
        WordprocessingMLUtil.InsertSectPrChildInOrder(sectPr, new XElement(W.type, new XAttribute(W.val, sectionStart)));
        return true;
    }
}
