using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Vba;

/// <summary>
/// VBA macro operations: list modules, view/import/delete code, run macros.
/// Requires VBA trust settings enabled in PowerPoint.
/// </summary>
[ServiceCategory("vba")]
[McpTool("vba", Title = "VBA Operations", Destructive = true, Category = "vba")]
public interface IVbaCommands
{
    /// <summary>
    /// List all VBA modules in the presentation.
    /// </summary>
    [ServiceAction("list")]
    VbaModuleListResult List(IPptBatch batch);

    /// <summary>
    /// View the code of a specific VBA module.
    /// </summary>
    [ServiceAction("view")]
    VbaModuleCodeResult View(IPptBatch batch, string moduleName);

    /// <summary>
    /// Import a new VBA module from code text.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="moduleName">Name for the new module</param>
    /// <param name="code">VBA code to import</param>
    /// <param name="moduleType">1=Standard, 2=ClassModule (default: 1)</param>
    [ServiceAction("import")]
    OperationResult Import(IPptBatch batch, string moduleName, string code, int moduleType);

    /// <summary>
    /// Delete a VBA module.
    /// </summary>
    [ServiceAction("delete")]
    OperationResult Delete(IPptBatch batch, string moduleName);

    /// <summary>
    /// Run a VBA macro by name.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="macroName">Fully qualified macro name (e.g., "Module1.MyMacro")</param>
    [ServiceAction("run")]
    OperationResult Run(IPptBatch batch, string macroName);
}
