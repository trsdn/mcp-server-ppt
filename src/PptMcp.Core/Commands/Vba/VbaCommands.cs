using System.Runtime.InteropServices;
using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Diagnostics;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Vba;

public class VbaCommands : IVbaCommands
{
    /// <summary>
    /// Acquires the VBA project, converting the default-Windows trust failure into
    /// actionable remediation (issue #130).
    ///
    /// The catch is filtered on the probe rather than on an HRESULT or message text.
    /// Matching the message would break on non-English Office, and matching a single
    /// HRESULT risks labelling unrelated COM failures as a trust problem. Consulting
    /// the setting itself answers the question directly.
    ///
    /// When trust IS enabled the filter is false, so the original exception propagates
    /// untouched to <c>batch.Execute()</c> exactly as before - this guard can only add
    /// information, never swallow a different fault.
    /// </summary>
    private static dynamic GetVbProject(dynamic pres)
    {
        try
        {
            return pres.VBProject;
        }
        catch (COMException ex) when (!VbaTrustProbe.IsTrustEnabled())
        {
            // Rule 1b: route a specific error, do not synthesise a result. The typed
            // exception propagates to batch.Execute(), which builds the
            // OperationResult at the correct layer.
            throw new VbaTrustException(ex);
        }
    }

    /// <summary>
    /// Reports whether VBA trust is enabled, without attempting a VBA operation.
    ///
    /// Exists so an agent can establish the precondition before it hits the failure,
    /// rather than having to provoke an error and interpret it. Rides the generated
    /// tool surface, so it is available on the CLI and the MCP server identically.
    /// </summary>
    public VbaTrustResult CheckTrust(IPptBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        // Deliberately not inside batch.Execute(): the answer comes from the registry,
        // not from the presentation, and it must remain obtainable when the COM call
        // that needs it would itself fail.
        var enabled = VbaTrustProbe.IsTrustEnabled();

        return new VbaTrustResult
        {
            Success = true,
            TrustEnabled = enabled,
            RegistryPath = VbaTrustProbe.DescribeRegistryPath(),
            Remediation = enabled ? string.Empty : VbaTrustProbe.RemediationText
        };
    }

    public VbaModuleListResult List(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = (dynamic)ctx.Presentation;
            dynamic vbProject = GetVbProject(pres);
            dynamic components = vbProject.VBComponents;
            try
            {
                int count = (int)components.Count;
                var result = new VbaModuleListResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath
                };

                for (int i = 1; i <= count; i++)
                {
                    dynamic comp = components.Item(i);
                    try
                    {
                        int modType = (int)comp.Type;
                        int lineCount = 0;
                        try { lineCount = (int)comp.CodeModule.CountOfLines; } catch { }

                        result.Modules.Add(new VbaModuleInfo
                        {
                            Name = comp.Name?.ToString() ?? "",
                            ModuleType = modType,
                            ModuleTypeName = GetModuleTypeName(modType),
                            LineCount = lineCount
                        });
                    }
                    finally
                    {
                        ComUtilities.Release(ref comp!);
                    }
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref components!);
                ComUtilities.Release(ref vbProject!);
            }
        });
    }

    public VbaModuleCodeResult View(IPptBatch batch, string moduleName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(moduleName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = (dynamic)ctx.Presentation;
            dynamic vbProject = GetVbProject(pres);
            dynamic components = vbProject.VBComponents;
            dynamic comp = components.Item(moduleName);
            dynamic? codeModule = null;
            try
            {
                codeModule = comp.CodeModule;
                int lineCount = (int)codeModule.CountOfLines;
                string code = lineCount > 0
                    ? codeModule.Lines(1, lineCount)?.ToString() ?? ""
                    : "";

                return new VbaModuleCodeResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    ModuleName = moduleName,
                    Code = code,
                    LineCount = lineCount
                };
            }
            finally
            {
                if (codeModule != null) ComUtilities.Release(ref codeModule!);
                ComUtilities.Release(ref comp!);
                ComUtilities.Release(ref components!);
                ComUtilities.Release(ref vbProject!);
            }
        });
    }

    public OperationResult Import(IPptBatch batch, string moduleName, string code, int moduleType)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(moduleName);
        ArgumentException.ThrowIfNullOrWhiteSpace(code);

        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = (dynamic)ctx.Presentation;
            dynamic vbProject = GetVbProject(pres);
            dynamic components = vbProject.VBComponents;
            dynamic? comp = null;
            dynamic? codeModule = null;
            try
            {
                // vbext_ct_StdModule = 1, vbext_ct_ClassModule = 2
                int vbType = moduleType == 2 ? 2 : 1;
                comp = components.Add(vbType);
                comp.Name = moduleName;
                codeModule = comp.CodeModule;
                codeModule.AddFromString(code);

                return new OperationResult
                {
                    Success = true,
                    Action = "import",
                    Message = $"Imported VBA module '{moduleName}'",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (codeModule != null) ComUtilities.Release(ref codeModule!);
                if (comp != null) ComUtilities.Release(ref comp!);
                ComUtilities.Release(ref components!);
                ComUtilities.Release(ref vbProject!);
            }
        });
    }

    public OperationResult Delete(IPptBatch batch, string moduleName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(moduleName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = (dynamic)ctx.Presentation;
            dynamic vbProject = GetVbProject(pres);
            dynamic components = vbProject.VBComponents;
            dynamic comp = components.Item(moduleName);
            try
            {
                components.Remove(comp);
                return new OperationResult
                {
                    Success = true,
                    Action = "delete",
                    Message = $"Deleted VBA module '{moduleName}'",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref comp!);
                ComUtilities.Release(ref components!);
                ComUtilities.Release(ref vbProject!);
            }
        });
    }

    public OperationResult Run(IPptBatch batch, string macroName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(macroName);

        return batch.Execute((ctx, ct) =>
        {
            // PowerPoint.Application.Run(macroName)
            dynamic app = ((dynamic)ctx.Presentation).Application;
            try
            {
                app.Run(macroName);
                return new OperationResult
                {
                    Success = true,
                    Action = "run",
                    Message = $"Executed macro '{macroName}'",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref app!);
            }
        });
    }

    private static string GetModuleTypeName(int moduleType) => moduleType switch
    {
        1 => "Standard",
        2 => "ClassModule",
        3 => "MSForm",
        100 => "Document",
        _ => $"Unknown({moduleType})"
    };
}
