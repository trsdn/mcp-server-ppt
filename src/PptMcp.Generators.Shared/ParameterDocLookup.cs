using System.Collections.Generic;
using System.Linq;
using Microsoft.CodeAnalysis;

namespace PptMcp.Generators.Common;

/// <summary>
/// Carries Core's <c>&lt;param&gt;</c> XML documentation across the compilation boundary.
///
/// <see cref="ServiceInfoExtractor"/> reads XML docs through
/// <c>ISymbol.GetDocumentationCommentXml()</c>, which only returns anything for *source*
/// symbols. `ServiceRegistryGenerator` runs inside PptMcp.Core and therefore sees the
/// docs, but the MCP and CLI generators run in their own projects where Core is a
/// metadata reference — and the C# compiler does not load XML documentation for
/// references. Every parameter description was silently dropped there, so 52 of 340 MCP
/// parameters shipped with an empty description and the remaining 288 carried nothing
/// but an auto-generated "(required …)" suffix (GitHub #128, Rule 18).
///
/// The fix is a generated type in Core whose fields are <c>const</c>. Constants *are*
/// written into metadata, so the downstream generators can read them back.
/// </summary>
public static class ParameterDocLookup
{
    /// <summary>Fully qualified name of the type emitted by ServiceRegistryGenerator.</summary>
    public const string TypeMetadataName = "PptMcp.Generated._ParameterDocs";

    /// <summary>
    /// Builds the lookup key for a parameter. Category and parameter are joined by a
    /// double underscore, which cannot occur inside either part.
    /// </summary>
    public static string Key(string categoryPascal, string parameterName)
        => $"{categoryPascal}__{Sanitize(parameterName)}";

    /// <summary>
    /// Reads the constants back out of the referenced Core assembly. Returns an empty
    /// dictionary when the type is absent, which keeps generation working during a
    /// clean build where Core has not been generated yet.
    /// </summary>
    public static Dictionary<string, string> Build(Compilation compilation)
    {
        var result = new Dictionary<string, string>();

        var type = compilation.GetTypeByMetadataName(TypeMetadataName);
        if (type is null)
            return result;

        foreach (var field in type.GetMembers().OfType<IFieldSymbol>())
        {
            if (field.HasConstantValue && field.ConstantValue is string value && value.Length > 0)
            {
                result[field.Name] = value;
            }
        }

        return result;
    }

    /// <summary>
    /// Fills in descriptions that the local compilation could not see. Never overwrites a
    /// description that is already present.
    ///
    /// Applied to the underlying <see cref="ParameterInfo"/> rather than to an aggregated
    /// parameter list, because <c>GetAllExposedParameters</c> rebuilds that list on every
    /// call — backfilling the aggregate would be discarded by the next caller.
    /// </summary>
    public static void Apply(IEnumerable<ServiceInfo> services, Dictionary<string, string> docs)
    {
        if (docs.Count == 0)
            return;

        foreach (var service in services)
        {
            foreach (var method in service.Methods)
            {
                foreach (var parameter in method.Parameters)
                {
                    if (!string.IsNullOrWhiteSpace(parameter.XmlDocDescription))
                        continue;

                    var exposedName = parameter.ExposedName ?? parameter.Name;
                    if (docs.TryGetValue(Key(service.CategoryPascal, exposedName), out var description))
                    {
                        parameter.XmlDocDescription = description;
                    }
                }
            }
        }
    }

    /// <summary>Maps a parameter name to a valid C# identifier suffix.</summary>
    public static string Sanitize(string name)
    {
        var chars = name.ToCharArray();
        for (var i = 0; i < chars.Length; i++)
        {
            if (!char.IsLetterOrDigit(chars[i]))
            {
                chars[i] = '_';
            }
        }

        return new string(chars);
    }
}
