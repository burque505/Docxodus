using System;
using System.Text.RegularExpressions;

namespace Docxodus.Internal;

/// <summary>
/// Recognizes the mail-merge field instructions — <c>MERGEFIELD</c>, <c>IF</c>, <c>=</c> formulas,
/// <c>FILLIN</c>, <c>ASK</c>, <c>NEXT</c>, <c>NEXTIF</c>, <c>SKIPIF</c>, <c>MERGEREC</c>,
/// <c>MERGESEQ</c> — for <see cref="WmlToHtmlConverterSettings.StampReferenceFields"/>: an editor
/// keeps such a field whole, as it does a cross-reference. The target is a MERGEFIELD's name, else empty.
/// </summary>
internal static class MergeFieldInstr
{
    /// <summary>A kind <see cref="TryParse"/> gives.</summary>
    internal static bool IsMergeKind(string kind) => kind is "MERGEFIELD" or "IF" or "=" or "FILLIN" or "ASK" or "NEXT"
        or "NEXTIF" or "SKIPIF" or "MERGEREC" or "MERGESEQ";

    internal static (string Kind, string Target, bool Hyperlink)? TryParse(string? instrText)
    {
        if (string.IsNullOrWhiteSpace(instrText)) return null;
        var t = instrText.TrimStart();
        if (t.StartsWith('=')) return ("=", "", false);
        var m = Regex.Match(t, @"^([A-Za-z]+)(?:\s+(""[^""]*""|[^\s\\]+))?");
        if (!m.Success) return null;
        var kind = m.Groups[1].Value.ToUpperInvariant();
        return kind switch
        {
            "MERGEFIELD" => m.Groups[2].Success ? (kind, m.Groups[2].Value.Trim('"'), false) : null,
            "IF" or "FILLIN" or "ASK" or "NEXT" or "NEXTIF" or "SKIPIF" or "MERGEREC" or "MERGESEQ" => (kind, "", false),
            _ => null,
        };
    }
}
