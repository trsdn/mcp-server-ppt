using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Design;

public class DesignCommands : IDesignCommands
{
    public DesignListResult List(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic designs = ((dynamic)ctx.Presentation).Designs;
            try
            {
                int count = (int)designs.Count;

                var result = new DesignListResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath
                };

                for (int i = 1; i <= count; i++)
                {
                    dynamic design = designs.Item(i);
                    try
                    {
                        int layoutCount = 0;
                        try
                        {
                            layoutCount = (int)design.SlideMaster.CustomLayouts.Count;
                        }
                        catch { }

                        result.Designs.Add(new DesignInfo
                        {
                            Index = i,
                            Name = design.Name?.ToString() ?? "",
                            LayoutCount = layoutCount
                        });
                    }
                    finally
                    {
                        ComUtilities.Release(ref design!);
                    }
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref designs!);
            }
        });
    }

    public OperationResult ApplyTheme(IPptBatch batch, string themePath)
    {
        return batch.Execute((ctx, ct) =>
        {
            if (!System.IO.File.Exists(themePath))
                throw new System.IO.FileNotFoundException($"Theme file not found: {themePath}");

            ((dynamic)ctx.Presentation).ApplyTheme(themePath);

            return new OperationResult
            {
                Success = true,
                Action = "apply-theme",
                Message = $"Applied theme from '{System.IO.Path.GetFileName(themePath)}'",
                FilePath = ctx.PresentationPath
            };
        });
    }

    public ThemeColorResult GetColors(IPptBatch batch, int designIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic designs = ((dynamic)ctx.Presentation).Designs;
            int idx = designIndex <= 0 ? 1 : designIndex;
            dynamic design = designs.Item(idx);
            dynamic? slideMaster = null;
            dynamic? theme = null;
            dynamic? colorScheme = null;
            try
            {
                slideMaster = design.SlideMaster;
                theme = slideMaster.Theme;
                colorScheme = theme.ThemeColorScheme;

                var colors = new Dictionary<string, string>();
                // MsoThemeColorSchemeIndex: 1-12
                string[] colorNames = [
                    "Dark1", "Light1", "Dark2", "Light2",
                    "Accent1", "Accent2", "Accent3", "Accent4",
                    "Accent5", "Accent6", "Hyperlink", "FollowedHyperlink"
                ];

                for (int i = 1; i <= Math.Min(12, colorNames.Length); i++)
                {
                    try
                    {
                        dynamic colorItem = colorScheme.Colors(i);
                        int rgb = (int)colorItem.RGB;
                        // COM returns BGR, convert to #RRGGBB
                        int r = rgb & 0xFF;
                        int g = (rgb >> 8) & 0xFF;
                        int b = (rgb >> 16) & 0xFF;
                        colors[colorNames[i - 1]] = $"#{r:X2}{g:X2}{b:X2}";
                        ComUtilities.Release(ref colorItem!);
                    }
                    catch { }
                }

                return new ThemeColorResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    DesignName = design.Name?.ToString() ?? "",
                    Colors = colors
                };
            }
            finally
            {
                if (colorScheme != null) ComUtilities.Release(ref colorScheme!);
                if (theme != null) ComUtilities.Release(ref theme!);
                if (slideMaster != null) ComUtilities.Release(ref slideMaster!);
                ComUtilities.Release(ref design!);
                ComUtilities.Release(ref designs!);
            }
        });
    }

    public ColorSchemeListResult ListColorSchemes(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic colorSchemes = ((dynamic)ctx.Presentation).ColorSchemes;
            try
            {
                var result = new ColorSchemeListResult { Success = true, FilePath = ctx.PresentationPath };
                int count = (int)colorSchemes.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic cs = colorSchemes.Item(i);
                    try
                    {
                        var info = new ColorSchemeInfo { Index = i };
                        // RGBColor indices: 1-8 map to standard PowerPoint color roles
                        string[] roleNames = ["Background", "Text", "Shadow", "Title", "Fill", "Accent1", "Accent2", "Accent3"];
                        for (int c = 1; c <= Math.Min(8, roleNames.Length); c++)
                        {
                            try
                            {
                                int rgb = (int)cs.Colors(c).RGB;
                                int r = rgb & 0xFF;
                                int g = (rgb >> 8) & 0xFF;
                                int b = (rgb >> 16) & 0xFF;
                                info.Colors[roleNames[c - 1]] = $"#{r:X2}{g:X2}{b:X2}";
                            }
                            catch { }
                        }
                        result.ColorSchemes.Add(info);
                    }
                    finally
                    {
                        ComUtilities.Release(ref cs!);
                    }
                }
                return result;
            }
            finally
            {
                ComUtilities.Release(ref colorSchemes!);
            }
        });
    }
}
