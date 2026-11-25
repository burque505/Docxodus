// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Docxodus;

/// <summary>Creating and changing paragraph and character styles.</summary>
public sealed partial class DocxSession
{
    /// <summary>The children of <c>w:style</c>, in CT_Style order.</summary>
    private static readonly string[] StyleChildOrder =
    {
        "name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority",
        "semiHidden", "unhideWhenUsed", "qFormat", "locked", "personal", "personalCompose",
        "personalReply", "rsid", "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr",
    };

    /// <summary>
    /// Create a paragraph or character style, or change one: what Word's Create New Style and
    /// Modify Style dialogs write. Only the non-null fields of <paramref name="op"/> change. The
    /// result's <see cref="EditResult.StyleId"/> names the style. One undo step; not tracked (Word
    /// records no revision for a style definition). Every paragraph and run in the style shows the
    /// change, so a caller re-renders the whole document.
    /// </summary>
    public EditResult SetStyle(StyleDefinitionOp op)
    {
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        var main = _doc!.MainDocumentPart;
        if (main is null) return EditResult.Fail(EditErrorCode.InternalError, "no main document part", null);

        static EditResult Invalid(string message) =>
            EditResult.Fail(EditErrorCode.InvalidStyleDefinition, message, null);

        var root = main.StyleDefinitionsPart?.GetXDocument().Root;
        XElement? Find(string id) => root?.Elements(W.style).FirstOrDefault(s => (string?)s.Attribute(W.styleId) == id);

        bool create = string.IsNullOrEmpty(op.StyleId);
        var style = create ? null : Find(op.StyleId!);
        if (!create && style is null)
            return EditResult.Fail(EditErrorCode.UnknownStyle, $"no style '{op.StyleId}'", null);

        string type = create ? op.Type ?? "paragraph" : (string?)style!.Attribute(W.type) ?? "paragraph";
        if (type is not ("paragraph" or "character"))
            return Invalid($"only paragraph and character styles can be set, not '{type}'");
        if (op.Type is not null && op.Type != type)
            return Invalid($"the style is a {type} style; a style's type cannot change");

        var name = op.Name?.Trim();
        if (create && string.IsNullOrEmpty(name)) return Invalid("a new style needs a name");
        if (name is not null)
        {
            if (name.Length == 0) return Invalid("a style's name cannot be empty");
            var current = (string?)style?.Element(W.name)?.Attribute(W.val);
            bool renaming = !string.Equals(name, current, StringComparison.Ordinal);
            if (renaming && style is not null && style.Attribute(W.customStyle).ToBoolean() != true)
                return Invalid($"'{current}' is a built-in style; its name cannot change");
            if (renaming && NameTaken(root, name, style))
                return Invalid($"there is already a style named '{name}'");
        }

        if (!string.IsNullOrEmpty(op.BasedOn))
        {
            var basis = Find(op.BasedOn);
            if (basis is null) return EditResult.Fail(EditErrorCode.UnknownStyle, $"no style '{op.BasedOn}'", null);
            if (((string?)basis.Attribute(W.type) ?? "paragraph") != type)
                return Invalid($"a {type} style can only be based on another {type} style");
            // Walk up from the new basis: meeting this style again would make a loop.
            var seen = new System.Collections.Generic.HashSet<string>();
            for (var id = op.BasedOn; id is not null && seen.Add(id);
                 id = (string?)Find(id)?.Element(W.basedOn)?.Attribute(W.val))
                if (!create && id == op.StyleId) return Invalid($"'{op.BasedOn}' is based on this style");
        }
        if (!string.IsNullOrEmpty(op.Next))
        {
            if (type != "paragraph") return Invalid("only a paragraph style has a style for the following paragraph");
            var next = Find(op.Next);
            if (next is null && !(create && op.Next == UniqueStyleId(root, name!)))
                return EditResult.Fail(EditErrorCode.UnknownStyle, $"no style '{op.Next}'", null);
            if (next is not null && (string?)next.Attribute(W.type) != "paragraph")
                return Invalid("the style for the following paragraph must be a paragraph style");
        }
        if (op.Paragraph is { } paragraph)
        {
            if (type != "paragraph") return Invalid("a character style has no paragraph formatting");
            if (ParagraphFormatProblem(paragraph) is { } problem) return Invalid(problem);
        }
        if (op.Run is { } run)
        {
            if (run.Code is not null || run.RunStyle is not null)
                return Invalid("a style's character formatting cannot name a character style");
            // CT_Style's rPr has no highlight: Word highlights text, never a style.
            if (!string.IsNullOrEmpty(run.Highlight) && run.Highlight != "none")
                return Invalid("a style cannot carry a highlight");
            try { ApplyFormatToRun(new XElement(W.r), run); }
            catch (ArgumentException ex) { return Invalid(ex.Message); }
        }

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var stylesPart = main.StyleDefinitionsPart ?? main.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>();
            var xDoc = stylesPart.GetXDocument();
            if (xDoc.Root is null) xDoc.Add(new XElement(W.styles, new XAttribute(XNamespace.Xmlns + "w", W.w)));
            root = xDoc.Root!;

            string styleId;
            if (create)
            {
                styleId = UniqueStyleId(root, name!);
                style = new XElement(W.style,
                    new XAttribute(W.type, type),
                    new XAttribute(W.customStyle, "1"),
                    new XAttribute(W.styleId, styleId),
                    new XElement(W.name, new XAttribute(W.val, name!)));
                var last = root.Elements(W.style).LastOrDefault();
                if (last is null) root.Add(style); else last.AddAfterSelf(style);
                // Word's defaults for a new style: based on the default style of its type, followed
                // by itself, shown in the gallery.
                var basis = op.BasedOn ?? DefaultStyleId(root, type);
                if (!string.IsNullOrEmpty(basis))
                    SetStyleChild(style, new XElement(W.basedOn, new XAttribute(W.val, basis)));
                if (type == "paragraph" && op.Next is null)
                    SetStyleChild(style, new XElement(W.next, new XAttribute(W.val, styleId)));
                if (op.QuickFormat is null) SetStyleChild(style, new XElement(W.qFormat));
            }
            else
            {
                styleId = op.StyleId!;
                if (name is not null) style!.Element(W.name)!.SetAttributeValue(W.val, name);
            }

            if (op.BasedOn is not null)
            {
                style!.Elements(W.basedOn).Remove();
                if (op.BasedOn.Length > 0) SetStyleChild(style, new XElement(W.basedOn, new XAttribute(W.val, op.BasedOn)));
            }
            if (op.Next is not null)
            {
                style!.Elements(W.next).Remove();
                if (op.Next.Length > 0) SetStyleChild(style, new XElement(W.next, new XAttribute(W.val, op.Next)));
            }
            if (op.QuickFormat is { } quick)
            {
                style!.Elements(W.qFormat).Remove();
                if (quick)
                {
                    // A style in the gallery is not hidden from it.
                    style.Elements(W.hidden).Remove();
                    style.Elements(W.semiHidden).Remove();
                    style.Elements(W.unhideWhenUsed).Remove();
                    SetStyleChild(style, new XElement(W.qFormat));
                }
            }

            if (op.Paragraph is { } paragraphOp)
            {
                var pPr = style!.Element(W.pPr) ?? SetStyleChild(style, new XElement(W.pPr));
                // The probe is a paragraph in this very style: with a toggle's own element removed,
                // it resolves to what the style inherits.
                WriteParagraphFormat(pPr, paragraphOp, StyleProbe(styleId, type));
                if (!pPr.HasElements && !pPr.HasAttributes) pPr.Remove();
            }

            if (op.Run is { } runOp)
            {
                WriteStyleRunFormat(style!, styleId, type, runOp);
                // Word keeps a paragraph style and its linked character style ("Heading 1 Char")
                // alike: a change to one's character formatting is made to the other too.
                if (type == "paragraph" && (string?)style!.Element(W.link)?.Attribute(W.val) is { } linkId
                    && Find(linkId) is { } linked && (string?)linked.Attribute(W.type) == "character")
                    WriteStyleRunFormat(linked, linkId, "character", runOp);
            }

            stylesPart.PutXDocument();
            InvalidateProjectionCache();
            return new EditResult { Success = true, StyleId = styleId };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, null);
        }
    }

    /// <summary>
    /// Delete a custom paragraph or character style, as Word's Styles pane does. Paragraphs in a
    /// deleted paragraph style take the style it was based on (the default paragraph style when
    /// none); runs in a deleted character style take its base the same way. Styles based on it are
    /// re-based on its base, a following-style link to it is dropped, and a paragraph style's
    /// linked character style ("My Style Char") goes with it. Numbering levels tied to it lose the
    /// tie. Built-in and default styles are refused (<see cref="EditErrorCode.StyleNotDeletable"/>).
    /// One undo step; not tracked (like <see cref="SetStyle"/>). A caller re-renders the whole
    /// document.
    /// </summary>
    public EditResult DeleteStyle(string styleId)
    {
        ArgumentNullException.ThrowIfNull(styleId);
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        var main = _doc!.MainDocumentPart;
        var stylesPart = main?.StyleDefinitionsPart;
        var root = stylesPart?.GetXDocument().Root;
        XElement? Find(string id) => root?.Elements(W.style).FirstOrDefault(s => (string?)s.Attribute(W.styleId) == id);

        var style = Find(styleId);
        if (style is null) return EditResult.Fail(EditErrorCode.UnknownStyle, $"no style '{styleId}'", null);
        string type = (string?)style.Attribute(W.type) ?? "paragraph";
        string name = (string?)style.Element(W.name)?.Attribute(W.val) ?? styleId;
        if (type is not ("paragraph" or "character"))
            return EditResult.Fail(EditErrorCode.StyleNotDeletable, $"only paragraph and character styles can be deleted, not '{type}'", null);
        if (style.Attribute(W._default).ToBoolean() == true)
            return EditResult.Fail(EditErrorCode.StyleNotDeletable, $"'{name}' is the default {type} style", null);
        if (style.Attribute(W.customStyle).ToBoolean() != true)
            return EditResult.Fail(EditErrorCode.StyleNotDeletable, $"'{name}' is a built-in style; Word keeps it", null);

        // What goes: the style, and for a paragraph style its linked character style (when that
        // is custom and links back). A deleted character style's partner just loses the link.
        var doomed = new System.Collections.Generic.Dictionary<string, XElement> { [styleId] = style };
        if (type == "paragraph" && (string?)style.Element(W.link)?.Attribute(W.val) is { } linkId
            && Find(linkId) is { } linked && (string?)linked.Attribute(W.type) == "character"
            && linked.Attribute(W.customStyle).ToBoolean() == true && linked.Attribute(W._default).ToBoolean() != true
            && (string?)linked.Element(W.link)?.Attribute(W.val) == styleId)
            doomed[linkId] = linked;

        // Each doomed style's replacement: its nearest surviving base (null = the type's default,
        // written as no style reference at all).
        string? Replacement(string id)
        {
            var seen = new System.Collections.Generic.HashSet<string>();
            var current = doomed[id];
            while ((string?)current.Element(W.basedOn)?.Attribute(W.val) is { } basis && seen.Add(basis))
            {
                var next = Find(basis);
                if (next is null) return null;
                if (!doomed.ContainsKey(basis))
                    return next.Attribute(W._default).ToBoolean() == true ? null : basis;
                current = next;
            }
            return null;
        }
        var replacements = doomed.Keys.ToDictionary(id => id, Replacement);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            foreach (var part in EnumerateProjectedParts().Append(main!.NumberingDefinitionsPart).OfType<DocumentFormat.OpenXml.Packaging.OpenXmlPart>())
            {
                if (part == stylesPart) continue;
                var partRoot = part.GetXDocument().Root;
                if (partRoot is null) continue;
                bool changed = false;
                foreach (var reference in partRoot.Descendants().Where(e => e.Name == W.pStyle || e.Name == W.rStyle).ToList())
                {
                    if ((string?)reference.Attribute(W.val) is not { } id || !replacements.TryGetValue(id, out var to)) continue;
                    changed = true;
                    // A numbering level's tie to a paragraph style is dropped, not moved.
                    if (to is null || reference.Parent?.Name == W.lvl) Unreference(reference);
                    else reference.SetAttributeValue(W.val, to);
                }
                if (changed) part.PutXDocument();
            }

            foreach (var other in root!.Elements(W.style).Where(s => !doomed.ContainsValue(s)))
            {
                if ((string?)other.Element(W.basedOn)?.Attribute(W.val) is { } basis && replacements.TryGetValue(basis, out var to))
                {
                    other.Elements(W.basedOn).Remove();
                    var fallback = to ?? DefaultStyleId(root, (string?)other.Attribute(W.type) ?? "paragraph");
                    if (!string.IsNullOrEmpty(fallback) && fallback != (string?)other.Attribute(W.styleId))
                        SetStyleChild(other, new XElement(W.basedOn, new XAttribute(W.val, fallback)));
                }
                if ((string?)other.Element(W.next)?.Attribute(W.val) is { } nextId && doomed.ContainsKey(nextId))
                    other.Elements(W.next).Remove();
                if ((string?)other.Element(W.link)?.Attribute(W.val) is { } link && doomed.ContainsKey(link))
                    other.Elements(W.link).Remove();
            }
            foreach (var gone in doomed.Values) gone.Remove();

            stylesPart!.PutXDocument();
            InvalidateProjectionCache();
            return new EditResult { Success = true, StyleId = styleId };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, null);
        }

        // Remove a style reference, then its pPr/rPr when nothing is left in it.
        static void Unreference(XElement reference)
        {
            var parent = reference.Parent;
            reference.Remove();
            if (parent is not null && (parent.Name == W.rPr || parent.Name == W.pPr)
                && !parent.HasElements && !parent.HasAttributes && parent.Parent?.Name != W.lvl)
                parent.Remove();
        }
    }

    /// <summary>Write <paramref name="op"/> into a style's <c>w:rPr</c> as into a run's, then make
    /// an off stick: where the style still inherits the property on, write an explicit off.</summary>
    private void WriteStyleRunFormat(XElement style, string styleId, string type, FormatOp op)
    {
        var rPr = style.Element(W.rPr) ?? new XElement(W.rPr);
        if (rPr.Parent is not null) rPr.Remove();
        var carrier = new XElement(W.r, rPr);
        ApplyFormatToRun(carrier, op);
        rPr = carrier.Element(W.rPr)!;
        rPr.Remove();
        SetStyleChild(style, rPr);

        void Off(bool? value, XName name, XElement explicitOff)
        {
            if (value is not false) return;
            var probe = StyleProbe(styleId, type);
            var inherited = FormattingAssembler.ResolveEffectiveRunProperties(_doc!, probe.Element(W.r)!).Element(name);
            bool on = name == W.u
                ? inherited is not null && (string?)inherited.Attribute(W.val) is not (null or "none")
                : inherited is not null && inherited.Attribute(W.val).ToBoolean() != false;
            if (on) WordprocessingMLUtil.InsertRPrChildInOrder(rPr, explicitOff);
        }
        Off(op.Bold, W.b, new XElement(W.b, new XAttribute(W.val, "0")));
        Off(op.Italic, W.i, new XElement(W.i, new XAttribute(W.val, "0")));
        Off(op.Strike, W.strike, new XElement(W.strike, new XAttribute(W.val, "0")));
        Off(op.Underline, W.u, new XElement(W.u, new XAttribute(W.val, "none")));
        Off(op.Caps, W.caps, new XElement(W.caps, new XAttribute(W.val, "0")));
        Off(op.SmallCaps, W.smallCaps, new XElement(W.smallCaps, new XAttribute(W.val, "0")));
        if (!rPr.HasElements && !rPr.HasAttributes) rPr.Remove();
    }

    /// <summary>A detached paragraph in the style (or, for a character style, holding a run in it),
    /// to resolve what the style gives.</summary>
    private static XElement StyleProbe(string styleId, string type) =>
        new(W.p,
            new XElement(W.pPr, type == "paragraph" ? new XElement(W.pStyle, new XAttribute(W.val, styleId)) : null),
            new XElement(W.r,
                new XElement(W.rPr, type == "character" ? new XElement(W.rStyle, new XAttribute(W.val, styleId)) : null),
                new XElement(W.t, "x")));

    /// <summary>Put <paramref name="child"/> in <paramref name="style"/> at its CT_Style place,
    /// replacing one of the same name.</summary>
    private static XElement SetStyleChild(XElement style, XElement child)
    {
        style.Elements(child.Name).Remove();
        int index = Array.IndexOf(StyleChildOrder, child.Name.LocalName);
        var before = style.Elements().FirstOrDefault(e => Array.IndexOf(StyleChildOrder, e.Name.LocalName) is var i
            && (i > index || i < 0));
        if (before is null) style.Add(child); else before.AddBeforeSelf(child);
        return child;
    }

    /// <summary>Whether a style other than <paramref name="except"/> has <paramref name="name"/>,
    /// or Word keeps the name for a built-in style (a latent style); case does not matter.</summary>
    private static bool NameTaken(XElement? root, string name, XElement? except)
    {
        if (root is null) return false;
        bool Same(string? other) => string.Equals(other, name, StringComparison.OrdinalIgnoreCase);
        return root.Elements(W.style).Any(s => s != except && Same((string?)s.Element(W.name)?.Attribute(W.val)))
            || (except is null && root.Element(W.latentStyles)?.Elements(W.lsdException)
                .Any(e => Same((string?)e.Attribute(W.name))) == true);
    }

    /// <summary>An id made from <paramref name="name"/> as Word makes one (its ASCII letters and
    /// digits), numbered when another style has it.</summary>
    private static string UniqueStyleId(XElement? root, string name)
    {
        var sb = new StringBuilder();
        foreach (var c in name) if (c is (>= 'A' and <= 'Z') or (>= 'a' and <= 'z') or (>= '0' and <= '9')) sb.Append(c);
        var stem = sb.Length > 0 ? sb.ToString() : "Style";
        bool Taken(string id) => root?.Elements(W.style).Any(s =>
            string.Equals((string?)s.Attribute(W.styleId), id, StringComparison.OrdinalIgnoreCase)) == true;
        if (!Taken(stem)) return stem;
        for (int n = 1; ; n++) if (!Taken(stem + n)) return stem + n;
    }

    private static string? DefaultStyleId(XElement root, string type) =>
        (string?)root.Elements(W.style)
            .FirstOrDefault(s => (string?)s.Attribute(W.type) == type && s.Attribute(W._default).ToBoolean() == true)?
            .Attribute(W.styleId);
}
