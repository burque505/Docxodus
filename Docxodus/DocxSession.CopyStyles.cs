// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus;

/// <summary>Copying styles from another document or template (Word's Organizer, from a file into
/// this document).</summary>
public sealed partial class DocxSession
{
    /// <summary>
    /// The styles of another document or template (<c>.docx</c>, <c>.dotx</c>, <c>.docm</c>,
    /// <c>.dotm</c>), as <see cref="ListStyles"/> lists this document's: what
    /// <see cref="CopyStyles"/> can copy from it. Throws <see cref="ArgumentException"/> when the
    /// bytes are not a Word package.
    /// </summary>
    public static IReadOnlyList<StyleInfo> ListStylesIn(byte[] documentBytes)
    {
        using var source = OpenStyleSource(documentBytes);
        return Internal.FormattingIntrospectionOps.ListStyles(source.Document);
    }

    /// <summary>
    /// Copy the styles <paramref name="styleIds"/> (ids in <paramref name="documentBytes"/>, as
    /// <see cref="ListStylesIn"/> gives them) into this document, as Word's Organizer copies them.
    /// One undo step; not tracked (Word records no revision for a style definition).
    /// <list type="bullet">
    /// <item>A style is matched to this document's by name and type, case ignored. A match is
    /// replaced (it keeps its id, so its paragraphs and runs take the copy's look); otherwise the
    /// style is added, under its own id unless this document has that id already.</item>
    /// <item>A based-on or linked style (and a list style a list in the style refers to) that this
    /// document lacks comes along; one this document has is used as it is, never replaced.</item>
    /// <item>A style for the following paragraph that this document lacks is dropped (the
    /// paragraph is followed by its own style).</item>
    /// <item>List numbering in a style is copied as a list definition of its own.</item>
    /// <item>Only a replaced default style stays the default of its type: this document's defaults
    /// do not change.</item>
    /// </list>
    /// Every paragraph and run in a copied style shows the change, so a caller re-renders the whole
    /// document.
    /// </summary>
    public EditResult CopyStyles(byte[] documentBytes, IReadOnlyList<string> styleIds)
    {
        ArgumentNullException.ThrowIfNull(documentBytes);
        ArgumentNullException.ThrowIfNull(styleIds);
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        var main = _doc!.MainDocumentPart;
        if (main is null) return EditResult.Fail(EditErrorCode.InternalError, "no main document part", null);
        if (styleIds.Count == 0)
            return EditResult.Fail(EditErrorCode.InvalidStyleDefinition, "no style to copy", null);

        StyleSource source;
        try { source = OpenStyleSource(documentBytes); }
        catch (ArgumentException ex) { return EditResult.Fail(EditErrorCode.InvalidStyleSource, ex.Message, null); }

        using (source)
        {
            var sourceRoot = source.Document.MainDocumentPart?.StyleDefinitionsPart?.GetXDocument().Root;
            var sourceStyles = sourceRoot?.Elements(W.style)
                .Where(s => !string.IsNullOrEmpty((string?)s.Attribute(W.styleId)))
                .GroupBy(s => (string)s.Attribute(W.styleId)!, StringComparer.Ordinal)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.Ordinal)
                ?? new Dictionary<string, XElement>(StringComparer.Ordinal);
            foreach (var id in styleIds)
                if (id is null || !sourceStyles.ContainsKey(id))
                    return EditResult.Fail(EditErrorCode.UnknownStyle, $"the other document has no style '{id}'", null);

            _history.RecordPreOp(TakeSnapshot());
            try
            {
                var stylesPart = main.StyleDefinitionsPart ?? main.AddNewPart<StyleDefinitionsPart>();
                var xDoc = stylesPart.GetXDocument();
                if (xDoc.Root is null) xDoc.Add(new XElement(W.styles, new XAttribute(XNamespace.Xmlns + "w", W.w)));
                var root = xDoc.Root!;
                var copier = new StyleCopier(source.Document, sourceStyles, main, root);
                copier.Copy(styleIds);
                stylesPart.PutXDocument();
                copier.SaveNumbering();
                InvalidateProjectionCache();
                return new EditResult { Success = true, StyleId = copier.TargetIdOf(styleIds[0]) };
            }
            catch (Exception ex)
            {
                return FailInternal(ex, null);
            }
        }
    }

    private sealed class StyleSource : IDisposable
    {
        private readonly MemoryStream _stream;
        public WordprocessingDocument Document { get; }

        public StyleSource(MemoryStream stream, WordprocessingDocument document)
        {
            _stream = stream;
            Document = document;
        }

        public void Dispose()
        {
            Document.Dispose();
            _stream.Dispose();
        }
    }

    /// <summary>Open another package to read its styles: through the reading funnel every opened
    /// package takes (strict conversion, depth door), on a copy of the bytes.</summary>
    private static StyleSource OpenStyleSource(byte[] documentBytes)
    {
        ArgumentNullException.ThrowIfNull(documentBytes);
        if (documentBytes.Length == 0) throw new ArgumentException("the file is empty", nameof(documentBytes));
        var stream = new MemoryStream();
        try
        {
            var bytes = StrictOoxmlNormalizer.NormalizeForReading(documentBytes);
            stream.Write(bytes, 0, bytes.Length);
            stream.Position = 0;
            var document = WordprocessingDocument.Open(stream, isEditable: true);
            Internal.PartDepthGuard.MarkMeasured(document);
            if (document.MainDocumentPart is null)
            {
                document.Dispose();
                throw new ArgumentException("the file is not a Word document or template", nameof(documentBytes));
            }
            return new StyleSource(stream, document);
        }
        catch (Exception ex) when (ex is not ArgumentException)
        {
            stream.Dispose();
            throw new ArgumentException($"the file is not a Word document or template: {ex.Message}", nameof(documentBytes), ex);
        }
        catch
        {
            stream.Dispose();
            throw;
        }
    }

    /// <summary>One <see cref="CopyStyles"/> run: which source styles go where, under which ids,
    /// and the list definitions they bring.</summary>
    private sealed class StyleCopier
    {
        private readonly WordprocessingDocument _source;
        private readonly Dictionary<string, XElement> _sourceStyles;
        private readonly MainDocumentPart _main;
        private readonly XElement _root;
        // Source id -> id in this document, for every style the copy writes or uses.
        private readonly Dictionary<string, string> _ids = new(StringComparer.Ordinal);
        // Source id -> this document's style it replaces (null: added).
        private readonly Dictionary<string, XElement?> _replaced = new(StringComparer.Ordinal);
        private readonly List<string> _order = new();
        private readonly HashSet<string> _usedIds;
        private readonly Dictionary<string, string> _numIds = new(StringComparer.Ordinal);
        private XElement? _sourceNumbering;
        private XElement? _numbering;
        private NumberingDefinitionsPart? _numberingPart;

        public StyleCopier(WordprocessingDocument source, Dictionary<string, XElement> sourceStyles,
            MainDocumentPart main, XElement root)
        {
            _source = source;
            _sourceStyles = sourceStyles;
            _main = main;
            _root = root;
            _usedIds = root.Elements(W.style).Select(s => (string?)s.Attribute(W.styleId))
                .Where(id => id is not null).Select(id => id!).ToHashSet(StringComparer.Ordinal);
            _sourceNumbering = source.MainDocumentPart?.NumberingDefinitionsPart?.GetXDocument().Root;
        }

        public string? TargetIdOf(string sourceId) => _ids.TryGetValue(sourceId, out var id) ? id : null;

        public void Copy(IReadOnlyList<string> styleIds)
        {
            // The chosen styles first (each replaces its match), then what they need.
            var queue = new Queue<string>();
            foreach (var id in styleIds)
                if (!_ids.ContainsKey(id))
                {
                    Plan(id, replace: true);
                    queue.Enqueue(id);
                }
            while (queue.Count > 0)
            {
                var id = queue.Dequeue();
                foreach (var needed in Dependencies(_sourceStyles[id]))
                {
                    if (_ids.ContainsKey(needed) || !_sourceStyles.ContainsKey(needed)) continue;
                    // A style this document has is used as it is; only a missing one is copied.
                    if (Plan(needed, replace: false)) queue.Enqueue(needed);
                }
            }

            XElement? last = _root.Elements(W.style).LastOrDefault();
            foreach (var id in _order)
            {
                var copy = Rewrite(_sourceStyles[id], _ids[id], _replaced[id]);
                if (_replaced[id] is { } old) old.ReplaceWith(copy);
                else
                {
                    if (last is null)
                    {
                        // The styles' place: after w:docDefaults and w:latentStyles.
                        var before = _root.Elements().LastOrDefault(e => e.Name == W.docDefaults || e.Name == W.latentStyles);
                        if (before is null) _root.AddFirst(copy); else before.AddAfterSelf(copy);
                    }
                    else last.AddAfterSelf(copy);
                    last = copy;
                }
            }
        }

        /// <summary>Decide what happens to one source style. Returns whether it is written (and its
        /// own needs must be looked at).</summary>
        private bool Plan(string sourceId, bool replace)
        {
            var style = _sourceStyles[sourceId];
            var match = Match(style);
            if (match is not null && !replace)
            {
                _ids[sourceId] = (string)match.Attribute(W.styleId)!;
                return false;
            }
            string id;
            if (match is not null) id = (string)match.Attribute(W.styleId)!;
            else
            {
                id = _usedIds.Contains(sourceId) ? FreeId(NameOf(style)) : sourceId;
                _usedIds.Add(id);
            }
            _ids[sourceId] = id;
            _replaced[sourceId] = match;
            _order.Add(sourceId);
            return true;
        }

        private static string NameOf(XElement style) =>
            (string?)style.Element(W.name)?.Attribute(W.val) ?? (string)style.Attribute(W.styleId)!;

        private static string TypeOf(XElement style) => (string?)style.Attribute(W.type) ?? "paragraph";

        /// <summary>This document's style of the same name and type (case ignored).</summary>
        private XElement? Match(XElement sourceStyle)
        {
            var name = NameOf(sourceStyle);
            var type = TypeOf(sourceStyle);
            return _root.Elements(W.style).FirstOrDefault(s => TypeOf(s) == type
                && string.Equals(NameOf(s), name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>An id from the style's name as Word makes one, free in this document.</summary>
        private string FreeId(string name)
        {
            var sb = new StringBuilder();
            foreach (var c in name) if (c is (>= 'A' and <= 'Z') or (>= 'a' and <= 'z') or (>= '0' and <= '9')) sb.Append(c);
            var stem = sb.Length > 0 ? sb.ToString() : "Style";
            if (!_usedIds.Contains(stem)) return stem;
            for (int n = 1; ; n++)
                if (!_usedIds.Contains(stem + n)) return stem + n;
        }

        /// <summary>The source styles a style needs: its based-on and linked styles, and the list
        /// style its numbering refers to.</summary>
        private IEnumerable<string> Dependencies(XElement style)
        {
            foreach (var name in new[] { W.basedOn, W.link })
                if ((string?)style.Element(name)?.Attribute(W.val) is { Length: > 0 } id) yield return id;
            if (NumIdOf(style) is { } numId && SourceAbstractOf(numId) is { } abstractNum
                && (string?)abstractNum.Element(W.numStyleLink)?.Attribute(W.val) is { Length: > 0 } listStyle)
                yield return listStyle;
        }

        private static string? NumIdOf(XElement style) =>
            (string?)style.Element(W.pPr)?.Element(W.numPr)?.Element(W.numId)?.Attribute(W.val);

        private XElement? SourceAbstractOf(string numId)
        {
            var num = _sourceNumbering?.Elements(W.num).FirstOrDefault(n => (string?)n.Attribute(W.numId) == numId);
            var abstractId = (string?)num?.Element(W.abstractNumId)?.Attribute(W.val);
            return abstractId is null ? null
                : _sourceNumbering!.Elements(W.abstractNum).FirstOrDefault(a => (string?)a.Attribute(W.abstractNumId) == abstractId);
        }

        /// <summary>This document's id for a source style: one the copy writes or uses, else this
        /// document's style of that name and type; null when there is none.</summary>
        private string? Resolve(string sourceId)
        {
            if (_ids.TryGetValue(sourceId, out var id)) return id;
            if (!_sourceStyles.TryGetValue(sourceId, out var style)) return null;
            return (string?)Match(style)?.Attribute(W.styleId);
        }

        /// <summary>The copy of a source style as this document holds it.</summary>
        private XElement Rewrite(XElement sourceStyle, string id, XElement? replaced)
        {
            var copy = new XElement(sourceStyle);
            // Only WordprocessingML: the source's extensions (and any annotations of ours) stay behind.
            copy.DescendantsAndSelf().Attributes().Where(a => !a.IsNamespaceDeclaration && a.Name.Namespace != W.w
                && a.Name.Namespace != XNamespace.None).Remove();
            copy.Descendants().Where(d => d.Name.Namespace != W.w).Remove();
            copy.SetAttributeValue(W.styleId, id);
            // A copy is the default of its type only where it replaces the default.
            copy.SetAttributeValue(W._default, replaced?.Attribute(W._default) is { } isDefault ? isDefault.Value : null);

            foreach (var name in new[] { W.basedOn, W.link, W.next })
            {
                var element = copy.Element(name);
                if ((string?)element?.Attribute(W.val) is not { Length: > 0 } from) continue;
                var to = Resolve(from);
                if (to is null) element!.Remove();
                else element!.SetAttributeValue(W.val, to);
            }

            var numId = copy.Element(W.pPr)?.Element(W.numPr)?.Element(W.numId);
            if ((string?)numId?.Attribute(W.val) is { Length: > 0 } sourceNum && sourceNum != "0")
            {
                if (CopyList(sourceNum) is { } targetNum) numId!.SetAttributeValue(W.val, targetNum);
                else numId!.Parent!.Remove();
            }
            return copy;
        }

        /// <summary>Copy a source list (w:num and its w:abstractNum) as a list definition of this
        /// document's own; returns its numId, or null when the source has no such list.</summary>
        private string? CopyList(string sourceNumId)
        {
            if (_numIds.TryGetValue(sourceNumId, out var known)) return known;
            var num = _sourceNumbering?.Elements(W.num).FirstOrDefault(n => (string?)n.Attribute(W.numId) == sourceNumId);
            var abstractNum = SourceAbstractOf(sourceNumId);
            if (num is null || abstractNum is null) return null;

            var numbering = Numbering();
            int NextId(XName element, XName attribute) =>
                numbering.Elements(element).Select(e => int.TryParse((string?)e.Attribute(attribute), out var n) ? n : -1)
                    .DefaultIfEmpty(-1).Max() + 1;

            var newAbstract = new XElement(abstractNum);
            newAbstract.DescendantsAndSelf().Attributes().Where(a => !a.IsNamespaceDeclaration && a.Name.Namespace != W.w
                && a.Name.Namespace != XNamespace.None).Remove();
            newAbstract.Descendants().Where(d => d.Name.Namespace != W.w).Remove();
            string abstractId = NextId(W.abstractNum, W.abstractNumId).ToString(System.Globalization.CultureInfo.InvariantCulture);
            newAbstract.SetAttributeValue(W.abstractNumId, abstractId);
            // A list id of its own, so Word does not take it for one of this document's lists.
            var used = numbering.Elements(W.abstractNum).Select(a => (string?)a.Element(W.nsid)?.Attribute(W.val))
                .Where(v => v is not null).ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (newAbstract.Element(W.nsid) is { } nsid && used.Contains((string?)nsid.Attribute(W.val) ?? ""))
                nsid.SetAttributeValue(W.val, FreeNsid(used));
            // Picture bullets would need their pictures: the level's own bullet text stays.
            newAbstract.Descendants(W.lvlPicBulletId).Remove();
            foreach (var link in new[] { W.styleLink, W.numStyleLink })
                if (newAbstract.Element(link) is { } element)
                {
                    var to = (string?)element.Attribute(W.val) is { Length: > 0 } from ? Resolve(from) : null;
                    if (to is null) element.Remove(); else element.SetAttributeValue(W.val, to);
                }
            foreach (var pStyle in newAbstract.Descendants(W.pStyle).ToList())
            {
                var to = (string?)pStyle.Attribute(W.val) is { Length: > 0 } from ? Resolve(from) : null;
                if (to is null) pStyle.Remove(); else pStyle.SetAttributeValue(W.val, to);
            }
            WordprocessingMLUtil.InsertNumberingChildInOrder(numbering, newAbstract);

            var newNum = new XElement(num);
            newNum.DescendantsAndSelf().Attributes().Where(a => !a.IsNamespaceDeclaration && a.Name.Namespace != W.w
                && a.Name.Namespace != XNamespace.None).Remove();
            newNum.Descendants().Where(d => d.Name.Namespace != W.w).Remove();
            string numId = Math.Max(1, NextId(W.num, W.numId)).ToString(System.Globalization.CultureInfo.InvariantCulture);
            newNum.SetAttributeValue(W.numId, numId);
            newNum.Element(W.abstractNumId)!.SetAttributeValue(W.val, abstractId);
            WordprocessingMLUtil.InsertNumberingChildInOrder(numbering, newNum);
            _numIds[sourceNumId] = numId;
            return numId;
        }

        private static string FreeNsid(HashSet<string?> used)
        {
            var random = new Random();
            while (true)
            {
                var value = random.Next(0x10000000, int.MaxValue).ToString("X8", System.Globalization.CultureInfo.InvariantCulture);
                if (!used.Contains(value)) return value;
            }
        }

        private XElement Numbering()
        {
            if (_numbering is not null) return _numbering;
            _numberingPart = _main.NumberingDefinitionsPart;
            if (_numberingPart is null)
            {
                _numberingPart = _main.AddNewPart<NumberingDefinitionsPart>();
                _numberingPart.PutXDocument(new XDocument(new XElement(W.numbering,
                    new XAttribute(XNamespace.Xmlns + "w", W.w.NamespaceName))));
            }
            var xDoc = _numberingPart.GetXDocument();
            if (xDoc.Root is null) xDoc.Add(new XElement(W.numbering, new XAttribute(XNamespace.Xmlns + "w", W.w.NamespaceName)));
            return _numbering = xDoc.Root!;
        }

        public void SaveNumbering() => _numberingPart?.PutXDocument();
    }
}
