using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Window;

public class WindowCommands : IWindowCommands
{
    public WindowInfoResult GetInfo(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = (dynamic)ctx.Presentation;
            dynamic app = pres.Application;
            try
            {
                int state = Convert.ToInt32(app.WindowState);
                return new WindowInfoResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    WindowState = state,
                    WindowStateName = GetWindowStateName(state),
                    Left = (float)app.Left,
                    Top = (float)app.Top,
                    Width = (float)app.Width,
                    Height = (float)app.Height
                };
            }
            finally
            {
                ComUtilities.Release(ref app!);
            }
        });
    }

    public OperationResult Minimize(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic app = ((dynamic)ctx.Presentation).Application;
            try
            {
                // ppWindowMinimized = 2
                app.WindowState = 2;
                return new OperationResult
                {
                    Success = true,
                    Action = "minimize",
                    Message = "PowerPoint window minimized",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref app!);
            }
        });
    }

    public OperationResult Restore(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic app = ((dynamic)ctx.Presentation).Application;
            try
            {
                // ppWindowNormal = 1
                app.WindowState = 1;
                return new OperationResult
                {
                    Success = true,
                    Action = "restore",
                    Message = "PowerPoint window restored",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref app!);
            }
        });
    }

    public OperationResult Maximize(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic app = ((dynamic)ctx.Presentation).Application;
            try
            {
                // ppWindowMaximized = 3
                app.WindowState = 3;
                return new OperationResult
                {
                    Success = true,
                    Action = "maximize",
                    Message = "PowerPoint window maximized",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref app!);
            }
        });
    }

    public OperationResult SetZoom(IPptBatch batch, int zoomPercent)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic app = ((dynamic)ctx.Presentation).Application;
            dynamic? window = null;
            try
            {
                window = app.ActiveWindow;
                dynamic view = window.View;
                try
                {
                    view.Zoom = zoomPercent;
                }
                finally
                {
                    ComUtilities.Release(ref view!);
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "set-zoom",
                    Message = $"Set zoom to {zoomPercent}%",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (window != null) ComUtilities.Release(ref window!);
                ComUtilities.Release(ref app!);
            }
        });
    }

    public OperationResult SetView(IPptBatch batch, int viewType)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic app = ((dynamic)ctx.Presentation).Application;
            dynamic? window = null;
            try
            {
                window = app.ActiveWindow;
                window.ViewType = viewType;
                return new OperationResult
                {
                    Success = true,
                    Action = "set-view",
                    Message = $"Set view to {GetViewTypeName(viewType)}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (window != null) ComUtilities.Release(ref window!);
                ComUtilities.Release(ref app!);
            }
        });
    }

    public OperationResult GetView(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic app = ((dynamic)ctx.Presentation).Application;
            dynamic? window = null;
            try
            {
                window = app.ActiveWindow;
                int viewType = Convert.ToInt32(window.ViewType);
                return new OperationResult
                {
                    Success = true,
                    Action = "get-view",
                    Message = $"Current view: {GetViewTypeName(viewType)} ({viewType})",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (window != null) ComUtilities.Release(ref window!);
                ComUtilities.Release(ref app!);
            }
        });
    }

    private static string GetWindowStateName(int state) => state switch
    {
        1 => "Normal",
        2 => "Minimized",
        3 => "Maximized",
        _ => $"Unknown({state})"
    };

    private static string GetViewTypeName(int viewType) => viewType switch
    {
        1 => "Normal",
        2 => "Outline",
        3 => "SlideSorter",
        4 => "NotesPage",
        5 => "SlideMaster",
        _ => $"Unknown({viewType})"
    };
}
