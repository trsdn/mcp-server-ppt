using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.DocumentProperty;

public class DocumentPropertyCommands : IDocumentPropertyCommands
{
    // Built-in property indices for BuiltinDocumentProperties.
    // These are raw ordinals into the OLE built-in property set, and a wrong one is
    // silent: writing and reading the same wrong index agrees with itself. The mapping
    // is pinned by DocumentPropertyRoundTripTests.BuiltInPropertyIndices_MapToTheExpectedNames,
    // which asserts each index's own Name against real PowerPoint.
    private const int PropTitle = 1;
    private const int PropSubject = 2;
    private const int PropAuthor = 3;
    private const int PropKeywords = 4;
    private const int PropComments = 5;
    private const int PropCompany = 21;
    private const int PropCategory = 18;

    public DocumentPropertyResult GetAll(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = (dynamic)ctx.Presentation;
            dynamic? builtIn = null;
            try
            {
                builtIn = pres.BuiltInDocumentProperties;
                return new DocumentPropertyResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    Title = GetProp(builtIn, PropTitle),
                    Subject = GetProp(builtIn, PropSubject),
                    Author = GetProp(builtIn, PropAuthor),
                    Keywords = GetProp(builtIn, PropKeywords),
                    Comments = GetProp(builtIn, PropComments),
                    Company = GetProp(builtIn, PropCompany),
                    Category = GetProp(builtIn, PropCategory)
                };
            }
            finally
            {
                if (builtIn != null) ComUtilities.Release(ref builtIn!);
            }
        });
    }

    public OperationResult SetAll(IPptBatch batch, string title, string subject, string author, string keywords, string comments, string company, string category)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = (dynamic)ctx.Presentation;
            dynamic? builtIn = null;
            try
            {
                builtIn = pres.BuiltInDocumentProperties;
                if (!string.IsNullOrEmpty(title)) SetProp(builtIn, PropTitle, title);
                if (!string.IsNullOrEmpty(subject)) SetProp(builtIn, PropSubject, subject);
                if (!string.IsNullOrEmpty(author)) SetProp(builtIn, PropAuthor, author);
                if (!string.IsNullOrEmpty(keywords)) SetProp(builtIn, PropKeywords, keywords);
                if (!string.IsNullOrEmpty(comments)) SetProp(builtIn, PropComments, comments);
                if (!string.IsNullOrEmpty(company)) SetProp(builtIn, PropCompany, company);
                if (!string.IsNullOrEmpty(category)) SetProp(builtIn, PropCategory, category);

                return new OperationResult
                {
                    Success = true,
                    Action = "set",
                    Message = "Updated document properties",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (builtIn != null) ComUtilities.Release(ref builtIn!);
            }
        });
    }

    private static string GetProp(dynamic props, int index)
    {
        dynamic? prop = null;
        try
        {
            prop = props.Item(index);
            return prop.Value?.ToString() ?? "";
        }
        finally
        {
            if (prop != null) ComUtilities.Release(ref prop!);
        }
    }

    private static void SetProp(dynamic props, int index, string value)
    {
        dynamic? prop = null;
        try
        {
            prop = props.Item(index);
            prop.Value = value;
        }
        finally
        {
            if (prop != null) ComUtilities.Release(ref prop!);
        }
    }

    public OperationResult GetCustom(IPptBatch batch, string propertyName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(propertyName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic customProps = pres.CustomDocumentProperties;
            dynamic? prop = null;
            try
            {
                prop = customProps.Item(propertyName);
                string value = prop.Value?.ToString() ?? "";

                return new OperationResult
                {
                    Success = true,
                    Action = "get-custom",
                    Message = $"{propertyName} = {value}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (prop != null) ComUtilities.Release(ref prop!);
                ComUtilities.Release(ref customProps!);
            }
        });
    }

    public OperationResult SetCustom(IPptBatch batch, string propertyName, string propertyValue)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(propertyName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic customProps = pres.CustomDocumentProperties;
            try
            {
                // COM has no TryGetItem: probing for an existing property means calling
                // Item and letting it throw. This catch is deliberate and stays - it is
                // the existence test itself, not error suppression. It is narrowed to the
                // probe alone so a failing *write* to an existing property still surfaces.
                bool exists = false;
                dynamic? existing = null;
                try
                {
                    existing = customProps.Item(propertyName);
                    exists = true;
                }
                catch { /* Property doesn't exist yet */ }

                try
                {
                    if (exists)
                    {
                        existing!.Value = propertyValue;
                    }
                    else
                    {
                        // Add new custom property (Type 4 = msoPropertyTypeString)
                        customProps.Add(propertyName, false, 4, propertyValue);
                    }
                }
                finally
                {
                    if (existing != null) ComUtilities.Release(ref existing!);
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "set-custom",
                    Message = $"Set custom property '{propertyName}' = '{propertyValue}'",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref customProps!);
            }
        });
    }
}
