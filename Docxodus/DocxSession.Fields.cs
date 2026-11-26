// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Docxodus.Internal;

namespace Docxodus;

public sealed partial class DocxSession
{
    private const int MaxFieldNesting = 8;

    /// <summary>
    /// Insert a <b>field</b> from its code, written as Word shows it with field codes on (Alt+F9):
    /// <c>MERGEFIELD Name \* Upper</c>, <c>IF { MERGEFIELD Guaranty } = True "Yes" "No"</c>,
    /// <c>= { MERGEFIELD Rate } * 100 \# "0.00%"</c>. Each <c>{ … }</c> pair becomes a field nested
    /// in the instruction, as Word writes one. One undo step.
    /// </summary>
    /// <remarks>
    /// <para>The fields are written as complex fields (<c>w:fldChar</c> runs), which can nest. A
    /// <c>MERGEFIELD</c> caches <c>«Name»</c>, as Word shows one before a merge; any other field
    /// caches <paramref name="cachedResult"/>, else a short <c>«…»</c> summary of its code (the
    /// field's own word and arguments, nested fields by their names, switches left out). Nothing
    /// is evaluated: a merge or Word's field update computes the result.</para>
    /// <para>Braces are always field delimiters, also inside quoted text, as in Word's field-code
    /// view. Unbalanced braces, an empty field or a nesting deeper than eight fail with
    /// <see cref="EditErrorCode.InvalidFieldCode"/>.</para>
    /// </remarks>
    public EditResult InsertField(string anchorId, int characterOffset, string code, string? cachedResult = null)
    {
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        if (ParseFieldCode(code, out var error) is not { } root)
            return EditResult.Fail(EditErrorCode.InvalidFieldCode, error!, anchorId);
        if (XmlText.ValidatePayload(code, anchorId) is { } codeError) return EditResult.Fail(codeError);
        if (cachedResult is not null && XmlText.ValidatePayload(cachedResult, anchorId) is { } resultError)
            return EditResult.Fail(resultError);

        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"InsertField requires a paragraph/heading/list-item anchor; got kind={target.Anchor.Kind}", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);
        var totalText = ParagraphText(element);
        if (characterOffset < 0 || characterOffset > totalText.Length)
            return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                $"offset {characterOffset} out of [0, {totalText.Length}]", anchorId);
        if (HasUnsupportedInlineInsertionBoundary(element, characterOffset))
            return EditResult.Fail(EditErrorCode.UnsupportedInlineBoundary,
                "InsertField offset falls inside a revision or unsupported inline container; " +
                "choose a boundary before or after that container", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            SplitRunsAtOffset(element, characterOffset, NextRevisionId);
            SplitInlineContainersAtOffset(element, characterOffset, NextRevisionId);
            var runs = FieldRuns(root, cachedResult);
            var insertion = new List<XElement>();
            if (_trackedChanges == TrackedChangeMode.RenderInline)
                insertion.Add(CreateRevisionEnvelope(W.ins, NewDistinctRevisionStamp(), runs.ToArray()));
            else insertion.AddRange(runs);
            foreach (var node in insertion) UnidHelper.AssignToSelfAndDescendants(node);
            InsertInlineAtOffset(element, characterOffset, insertion[0]);
            for (int i = 1; i < insertion.Count; i++) insertion[i - 1].AddAfterSelf(insertion[i]);
            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    /// <summary>A field's code: text pieces and nested fields, in order.</summary>
    private sealed class FieldCodeNode
    {
        public List<object> Parts { get; } = new();   // string | FieldCodeNode

        public string Flat => string.Concat(Parts.Select(p => p as string ?? " "));
        public string Word => FieldWordOf(Flat);
    }

    private static string FieldWordOf(string flat)
    {
        var t = flat.TrimStart();
        if (t.StartsWith('=')) return "=";
        var m = Regex.Match(t, @"^[A-Za-z]+");
        return m.Success ? m.Value.ToUpperInvariant() : "";
    }

    /// <summary>Parse <paramref name="code"/> (the outer braces optional) into a field tree, or
    /// null with <paramref name="error"/>.</summary>
    private static FieldCodeNode? ParseFieldCode(string? code, out string? error)
    {
        error = null;
        var text = (code ?? "").Trim();
        if (text.Length == 0) { error = "field code is empty"; return null; }
        // "{ X }" around the whole code is the field itself, not a nested one.
        if (text[0] == '{' && MatchingBrace(text, 0) == text.Length - 1) text = text[1..^1];
        var stack = new Stack<FieldCodeNode>();
        var root = new FieldCodeNode();
        var current = root;
        var piece = new StringBuilder();
        void Flush() { if (piece.Length > 0) { current.Parts.Add(piece.ToString()); piece.Clear(); } }
        foreach (var c in text)
        {
            if (c == '{')
            {
                Flush();
                if (stack.Count + 1 >= MaxFieldNesting) { error = "fields nest too deep"; return null; }
                var child = new FieldCodeNode();
                current.Parts.Add(child);
                stack.Push(current);
                current = child;
            }
            else if (c == '}')
            {
                if (stack.Count == 0) { error = "unbalanced '}' in field code"; return null; }
                Flush();
                if (current.Flat.Trim().Length == 0) { error = "empty nested field"; return null; }
                current = stack.Pop();
            }
            else piece.Append(c);
        }
        if (stack.Count > 0) { error = "unbalanced '{' in field code"; return null; }
        Flush();
        if (root.Flat.Trim().Length == 0) { error = "field code is empty"; return null; }
        if (root.Word.Length == 0) { error = "field code must start with a field name or '='"; return null; }
        return root;
    }

    private static int MatchingBrace(string text, int open)
    {
        int depth = 0;
        for (int i = open; i < text.Length; i++)
        {
            if (text[i] == '{') depth++;
            else if (text[i] == '}' && --depth == 0) return i;
        }
        return -1;
    }

    /// <summary>The text a field shows before a merge or an update.</summary>
    private static string FieldDisplay(FieldCodeNode node)
    {
        var flat = Summary(node);
        if (node.Word == "MERGEFIELD")
        {
            var name = Regex.Match(flat, @"^\s*MERGEFIELD\s+(""[^""]*""|\S+)", RegexOptions.IgnoreCase);
            if (name.Success) return "«" + name.Groups[1].Value.Trim('"') + "»";
        }
        // drop switches (and what follows them), collapse spaces
        var cut = Regex.Replace(flat, @"\s\\[*#@!a-zA-Z].*$", "", RegexOptions.Singleline);
        cut = Regex.Replace(cut, @"\s+", " ").Trim();
        if (cut.Length > 40) cut = cut[..39] + "…";
        return "«" + cut + "»";
    }

    /// <summary>The code with nested fields replaced by what they show (a merge field by its name).</summary>
    private static string Summary(FieldCodeNode node) => string.Concat(node.Parts.Select(p => p switch
    {
        string s => s,
        FieldCodeNode n => FieldDisplay(n).Trim('«', '»'),
        _ => "",
    }));

    private static List<XElement> FieldRuns(FieldCodeNode node, string? cachedResult, bool nested = false)
    {
        static XElement FldChar(string type) => new(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, type)));
        var runs = new List<XElement> { FldChar("begin") };
        foreach (var part in node.Parts)
        {
            if (part is string s)
                runs.Add(new XElement(W.r, new XElement(W.instrText, new XAttribute(XNamespace.Xml + "space", "preserve"), s)));
            else if (part is FieldCodeNode child)
                runs.AddRange(FieldRuns(child, null, nested: true));
        }
        // A nested field has no cached result: its text would count in the paragraph (and show)
        // beside the outer field's; Word computes it when the outer field is updated or merged.
        if (!nested)
        {
            runs.Add(FldChar("separate"));
            runs.Add(new XElement(W.r, TextContent(cachedResult ?? FieldDisplay(node), W.t)));
        }
        runs.Add(FldChar("end"));
        return runs;
    }
}
