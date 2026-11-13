#nullable enable

using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.Json;

namespace DocxodusNative;

/// <summary>
/// String-in, string-out entry point to the browser bridges (<c>DocxSessionBridge</c>,
/// <c>DocumentConverter</c>) for a desktop host whose page talks to it over a text channel —
/// a WebView2 host object, a synchronous request, a pipe.
///
/// <para>The page asks once for <see cref="Manifest"/> and builds, from it, an object with the
/// same functions as the WASM export it replaces. Each call is <see cref="Call"/> with the
/// method's name and its arguments as a JSON array. Arguments are what JSExport would have
/// received: numbers, strings, booleans, <c>null</c> (also for an argument the caller left out),
/// and a <c>Uint8Array</c> sent as a base64 string. The result comes back as text, read on the
/// page by the method's return type in the manifest: a string as is (null as null), a number or
/// boolean as its JSON text, bytes as base64, nothing for void. An exception propagates, as a
/// JSExport call's exception becomes a JS <c>Error</c>.</para>
///
/// <para>Calls must come from one thread at a time, as the browser's single-threaded runtime
/// guarantees for the WASM bridge (the session registry is not synchronized).</para>
/// </summary>
public static class NativeSessionBridge
{
    private static readonly Type[] Bridges =
    {
        typeof(DocxodusWasm.DocxSessionBridge),
        typeof(DocxodusWasm.DocumentConverter),
    };

    private static readonly Dictionary<string, MethodInfo> Methods = BuildMethods();

    private static Dictionary<string, MethodInfo> BuildMethods()
    {
        var methods = new Dictionary<string, MethodInfo>(StringComparer.Ordinal);
        foreach (var bridge in Bridges)
        {
            // Every public static method: the bridges' public surface is their [JSExport] surface, and the attribute
            // itself can't be read here (its assembly, System.Runtime.InteropServices.JavaScript, is browser-only).
            foreach (var m in bridge.GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly))
            {
                if (m.IsSpecialName || m.IsGenericMethodDefinition) continue;
                // JS sees one function per name; a second overload could never have been called.
                methods.TryAdd(bridge.Name + "." + m.Name, m);
            }
        }
        return methods;
    }

    /// <summary>
    /// JSON <c>{"Bridge.Method": {"p": [paramType, ...], "r": returnType}, ...}</c>; the types are
    /// <c>string</c>, <c>int</c>, <c>long</c>, <c>double</c>, <c>bool</c>, <c>bytes</c> or <c>void</c>
    /// (return only), with a trailing <c>?</c> when nullable.
    /// </summary>
    public static string Manifest()
    {
        var sb = new StringBuilder("{");
        foreach (var (name, m) in Methods)
        {
            if (sb.Length > 1) sb.Append(',');
            sb.Append(JsonSerializer.Serialize(name)).Append(":{\"p\":[");
            var ps = m.GetParameters();
            for (var i = 0; i < ps.Length; i++)
            {
                if (i > 0) sb.Append(',');
                sb.Append('"').Append(TypeName(ps[i].ParameterType, ps[i])).Append('"');
            }
            sb.Append("],\"r\":\"").Append(m.ReturnType == typeof(void) ? "void" : TypeName(m.ReturnType, m.ReturnParameter)).Append("\"}");
        }
        return sb.Append('}').ToString();
    }

    /// <summary>Calls <paramref name="method"/> (<c>"DocxSessionBridge.OpenSession"</c>) with the
    /// JSON array <paramref name="argsJson"/>; see the class remarks for the wire forms.</summary>
    public static string? Call(string method, string argsJson)
    {
        if (!Methods.TryGetValue(method, out var m))
            throw new MissingMethodException("No bridge method " + method);
        var ps = m.GetParameters();
        var args = new object?[ps.Length];
        using (var doc = JsonDocument.Parse(string.IsNullOrEmpty(argsJson) ? "[]" : argsJson))
        {
            var given = doc.RootElement;
            var count = given.ValueKind == JsonValueKind.Array ? given.GetArrayLength() : 0;
            for (var i = 0; i < ps.Length; i++)
                args[i] = i < count ? Convert(given[i], ps[i].ParameterType) : Missing(ps[i].ParameterType);
        }
        object? result;
        try
        {
            result = m.Invoke(null, args);
        }
        catch (TargetInvocationException e) when (e.InnerException is not null)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(e.InnerException).Throw();
            throw;
        }
        return result switch
        {
            null => null,
            string s => s,
            byte[] b => System.Convert.ToBase64String(b),
            bool f => f ? "true" : "false",
            IFormattable n => n.ToString(null, CultureInfo.InvariantCulture),
            _ => result.ToString(),
        };
    }

    private static object? Convert(JsonElement e, Type t)
    {
        if (e.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined) return Missing(t);
        var u = Nullable.GetUnderlyingType(t) ?? t;
        if (u == typeof(string)) return e.ValueKind == JsonValueKind.String ? e.GetString() : e.GetRawText();
        if (u == typeof(byte[])) return System.Convert.FromBase64String(e.GetString() ?? "");
        if (u == typeof(bool)) return e.ValueKind == JsonValueKind.True;
        if (u == typeof(double)) return e.GetDouble();
        if (u == typeof(float)) return (float)e.GetDouble();
        // JS numbers: an integer parameter takes the value's integer part, as JSExport's marshaller does.
        if (u == typeof(int)) return e.TryGetInt32(out var i) ? i : (int)e.GetDouble();
        if (u == typeof(long)) return e.TryGetInt64(out var l) ? l : (long)e.GetDouble();
        throw new NotSupportedException("Bridge parameter type " + t.Name);
    }

    /// <summary>A left-out or null argument: null where the parameter allows it, else the type's default.</summary>
    private static object? Missing(Type t) => t.IsValueType && Nullable.GetUnderlyingType(t) is null ? Activator.CreateInstance(t) : null;

    private static string TypeName(Type t, ParameterInfo p)
    {
        var u = Nullable.GetUnderlyingType(t);
        var nullable = u is not null || (!t.IsValueType && new NullabilityInfoContext().Create(p).ReadState == NullabilityState.Nullable);
        var name = (u ?? t) switch
        {
            var x when x == typeof(string) => "string",
            var x when x == typeof(int) => "int",
            var x when x == typeof(long) => "long",
            var x when x == typeof(double) || x == typeof(float) => "double",
            var x when x == typeof(bool) => "bool",
            var x when x == typeof(byte[]) => "bytes",
            var x => x.Name,
        };
        return nullable ? name + "?" : name;
    }
}
