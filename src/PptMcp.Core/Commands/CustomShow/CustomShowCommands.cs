using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.CustomShow;

public class CustomShowCommands : ICustomShowCommands
{
    public CustomShowListResult List(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic shows = pres.SlideShowSettings.NamedSlideShows;
            try
            {
                var result = new CustomShowListResult { Success = true, FilePath = ctx.PresentationPath };
                int count = (int)shows.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic show = shows.Item(i);
                    try
                    {
                        var info = new CustomShowInfo
                        {
                            Index = i,
                            Name = show.Name?.ToString() ?? "",
                            SlideCount = (int)show.Count
                        };
                        // Get slide IDs
                        for (int s = 1; s <= info.SlideCount; s++)
                        {
                            try { info.SlideIds.Add((int)show.SlideIDs(s)); } catch { }
                        }
                        result.Shows.Add(info);
                    }
                    finally
                    {
                        ComUtilities.Release(ref show!);
                    }
                }
                return result;
            }
            finally
            {
                ComUtilities.Release(ref shows!);
            }
        });
    }

    public OperationResult Create(IPptBatch batch, string showName, string slideIndices)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(showName);
        ArgumentException.ThrowIfNullOrWhiteSpace(slideIndices);

        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic shows = pres.SlideShowSettings.NamedSlideShows;
            try
            {
                int[] indices = slideIndices.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    .Select(s => int.Parse(s, System.Globalization.CultureInfo.InvariantCulture))
                    .ToArray();

                // Build array of slide IDs from indices
                dynamic slides = pres.Slides;
                try
                {
                    int[] slideIds = new int[indices.Length];
                    for (int i = 0; i < indices.Length; i++)
                    {
                        dynamic slide = slides.Item(indices[i]);
                        slideIds[i] = (int)slide.SlideID;
                        ComUtilities.Release(ref slide!);
                    }

                    shows.Add(showName, slideIds);
                }
                finally
                {
                    ComUtilities.Release(ref slides!);
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "create",
                    Message = $"Created custom show '{showName}' with {indices.Length} slide(s)",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref shows!);
            }
        });
    }

    public OperationResult Delete(IPptBatch batch, string showName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(showName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic shows = pres.SlideShowSettings.NamedSlideShows;
            try
            {
                dynamic show = shows.Item(showName);
                try
                {
                    show.Delete();
                }
                finally
                {
                    ComUtilities.Release(ref show!);
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "delete",
                    Message = $"Deleted custom show '{showName}'",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref shows!);
            }
        });
    }
}
