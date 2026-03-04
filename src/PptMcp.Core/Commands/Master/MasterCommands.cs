using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Master;

public class MasterCommands : IMasterCommands
{
    public MasterListResult List(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            var result = new MasterListResult { Success = true, FilePath = ctx.PresentationPath };
            dynamic pres = ctx.Presentation;
            dynamic masters = pres.SlideMasters;
            try
            {
                int masterCount = (int)masters.Count;

                for (int m = 1; m <= masterCount; m++)
                {
                    dynamic master = masters.Item(m);
                    try
                    {
                        var masterInfo = new MasterInfo
                        {
                            Name = master.Name?.ToString() ?? $"Master {m}"
                        };

                        dynamic layouts = master.CustomLayouts;
                        try
                        {
                            int layoutCount = (int)layouts.Count;
                            for (int l = 1; l <= layoutCount; l++)
                            {
                                dynamic layout = layouts.Item(l);
                                try
                                {
                                    masterInfo.Layouts.Add(new LayoutInfo
                                    {
                                        Name = layout.Name?.ToString() ?? $"Layout {l}",
                                        Index = l
                                    });
                                }
                                finally
                                {
                                    ComUtilities.Release(ref layout!);
                                }
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref layouts!);
                        }

                        result.Masters.Add(masterInfo);
                    }
                    finally
                    {
                        ComUtilities.Release(ref master!);
                    }
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref masters!);
            }
        });
    }
}
