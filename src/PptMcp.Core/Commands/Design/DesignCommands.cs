using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Design;

public partial class DesignCommands : IDesignCommands
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
                            layoutCount = ComUtilities.GetCustomLayoutCount(design);
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

    /// <summary>
    /// The twelve MsoThemeColorSchemeIndex roles, in index order.
    /// </summary>
    private static readonly string[] ThemeColorRoles =
    [
        "Dark1", "Light1", "Dark2", "Light2",
        "Accent1", "Accent2", "Accent3", "Accent4",
        "Accent5", "Accent6", "Hyperlink", "FollowedHyperlink"
    ];

    /// <summary>
    /// Reads one design's theme colour palette. Shared by <c>get-colors</c> and
    /// <c>list-color-schemes</c> so the two cannot report different palettes for the same design.
    /// </summary>
    /// <param name="design">A live Designs.Item(i) proxy. Ownership stays with the caller.</param>
    private static Dictionary<string, string> ReadThemeColors(dynamic design)
    {
        dynamic? slideMaster = null;
        dynamic? theme = null;
        dynamic? colorScheme = null;
        try
        {
            slideMaster = design.SlideMaster;
            theme = slideMaster.Theme;
            colorScheme = theme.ThemeColorScheme;

            var colors = new Dictionary<string, string>();

            for (int i = 1; i <= ThemeColorRoles.Length; i++)
            {
                dynamic colorItem = colorScheme.Colors(i);
                try
                {
                    // PowerPoint stores colours as 0x00BBGGRR, not #RRGGBB.
                    colors[ThemeColorRoles[i - 1]] = ComUtilities.FormatOleColorAsHex((int)colorItem.RGB);
                }
                finally
                {
                    ComUtilities.Release(ref colorItem!);
                }
            }

            return colors;
        }
        finally
        {
            if (colorScheme != null) ComUtilities.Release(ref colorScheme!);
            if (theme != null) ComUtilities.Release(ref theme!);
            if (slideMaster != null) ComUtilities.Release(ref slideMaster!);
        }
    }

    public ThemeColorResult GetColors(IPptBatch batch, int designIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic designs = ((dynamic)ctx.Presentation).Designs;
            int idx = designIndex <= 0 ? 1 : designIndex;
            dynamic design = designs.Item(idx);
            try
            {
                return new ThemeColorResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    DesignName = design.Name?.ToString() ?? "",
                    Colors = ReadThemeColors(design)
                };
            }
            finally
            {
                ComUtilities.Release(ref design!);
                ComUtilities.Release(ref designs!);
            }
        });
    }

    public ColorSchemeListResult ListColorSchemes(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            // Reads the theme colour scheme of every design, not Presentation.ColorSchemes.
            //
            // ColorSchemes is the pre-2007 API that Office themes replaced. It is empty for every
            // OOXML presentation, so this action used to return Success with an empty list on any
            // modern .pptx - indistinguishable from a genuine "this deck has no colour schemes",
            // while get-colors returned a full twelve-role palette for the same file (issue #174).
            //
            // A design is the modern unit that owns a palette, so one entry per design is the
            // faithful answer. Index addresses the same design get-colors takes.
            dynamic designs = ((dynamic)ctx.Presentation).Designs;
            try
            {
                var result = new ColorSchemeListResult { Success = true, FilePath = ctx.PresentationPath };
                int count = (int)designs.Count;

                for (int i = 1; i <= count; i++)
                {
                    dynamic design = designs.Item(i);
                    try
                    {
                        result.ColorSchemes.Add(new ColorSchemeInfo
                        {
                            Index = i,
                            DesignName = design.Name?.ToString() ?? "",
                            Colors = ReadThemeColors(design)
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

    public ThemeFontResult GetThemeFonts(IPptBatch batch, int designIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic designs = ((dynamic)ctx.Presentation).Designs;
            int idx = designIndex <= 0 ? 1 : designIndex;
            dynamic design = designs.Item(idx);
            dynamic? slideMaster = null;
            dynamic? theme = null;
            dynamic? fontScheme = null;
            dynamic? majorFont = null;
            dynamic? minorFont = null;
            try
            {
                slideMaster = design.SlideMaster;
                theme = slideMaster.Theme;
                fontScheme = theme.ThemeFontScheme;
                majorFont = fontScheme.MajorFont;
                minorFont = fontScheme.MinorFont;

                // Item(1) = Latin font
                string headingFont = majorFont.Item(1).Name?.ToString() ?? "";
                string bodyFont = minorFont.Item(1).Name?.ToString() ?? "";

                return new ThemeFontResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    DesignName = design.Name?.ToString() ?? "",
                    HeadingFont = headingFont,
                    BodyFont = bodyFont
                };
            }
            finally
            {
                if (minorFont != null) ComUtilities.Release(ref minorFont!);
                if (majorFont != null) ComUtilities.Release(ref majorFont!);
                if (fontScheme != null) ComUtilities.Release(ref fontScheme!);
                if (theme != null) ComUtilities.Release(ref theme!);
                if (slideMaster != null) ComUtilities.Release(ref slideMaster!);
                ComUtilities.Release(ref design!);
                ComUtilities.Release(ref designs!);
            }
        });
    }
}
