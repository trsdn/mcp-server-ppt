using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Hyperlink;

public class HyperlinkCommands : IHyperlinkCommands
{
    public OperationResult Add(IPptBatch batch, int slideIndex, string shapeName, string address, string? subAddress = null, string? screenTip = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);
        ArgumentException.ThrowIfNullOrWhiteSpace(address);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? actionSettings = null;
            dynamic? actionSetting = null;
            dynamic? hyperlink = null;
            try
            {
                actionSettings = shape.ActionSettings;
                // ppMouseClick = 1
                actionSetting = actionSettings.Item(1);
                // ppActionHyperlink = 7
                actionSetting.Action = 7;
                hyperlink = actionSetting.Hyperlink;
                hyperlink.Address = address;
                hyperlink.SubAddress = subAddress;
                if (!string.IsNullOrEmpty(screenTip))
                    hyperlink.ScreenTip = screenTip;

                return new OperationResult
                {
                    Success = true,
                    Action = "add",
                    Message = $"Added hyperlink '{address}' to shape '{shapeName}' on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (hyperlink != null) ComUtilities.Release(ref hyperlink!);
                if (actionSetting != null) ComUtilities.Release(ref actionSetting!);
                if (actionSettings != null) ComUtilities.Release(ref actionSettings!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public HyperlinkResult Read(IPptBatch batch, int slideIndex, string shapeName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? actionSettings = null;
            dynamic? actionSetting = null;
            dynamic? hyperlink = null;
            try
            {
                actionSettings = shape.ActionSettings;
                actionSetting = actionSettings.Item(1); // ppMouseClick = 1
                int action = Convert.ToInt32(actionSetting.Action);

                string address = "";
                string subAddress = "";
                string screenTip = "";
                bool hasHyperlink = action == 7; // ppActionHyperlink

                if (hasHyperlink)
                {
                    hyperlink = actionSetting.Hyperlink;
                    address = hyperlink.Address?.ToString() ?? "";
                    subAddress = hyperlink.SubAddress?.ToString() ?? "";
                    try { screenTip = hyperlink.ScreenTip?.ToString() ?? ""; } catch { }
                }

                return new HyperlinkResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    SlideIndex = slideIndex,
                    ShapeName = shapeName,
                    HasHyperlink = hasHyperlink,
                    Address = address,
                    SubAddress = subAddress,
                    ScreenTip = screenTip
                };
            }
            finally
            {
                if (hyperlink != null) ComUtilities.Release(ref hyperlink!);
                if (actionSetting != null) ComUtilities.Release(ref actionSetting!);
                if (actionSettings != null) ComUtilities.Release(ref actionSettings!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult Remove(IPptBatch batch, int slideIndex, string shapeName)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? actionSettings = null;
            dynamic? actionSetting = null;
            try
            {
                actionSettings = shape.ActionSettings;
                actionSetting = actionSettings.Item(1); // ppMouseClick = 1
                // ppActionNone = 0
                actionSetting.Action = 0;

                return new OperationResult
                {
                    Success = true,
                    Action = "remove",
                    Message = $"Removed hyperlink from shape '{shapeName}' on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (actionSetting != null) ComUtilities.Release(ref actionSetting!);
                if (actionSettings != null) ComUtilities.Release(ref actionSettings!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public HyperlinkListResult List(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = (dynamic)ctx.Presentation;
            dynamic slides = pres.Slides;
            try
            {
                int slideCount = Convert.ToInt32(slides.Count);

                var result = new HyperlinkListResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath
                };

                int globalIndex = 1;
                for (int si = 1; si <= slideCount; si++)
                {
                    dynamic? slide = null;
                    dynamic? shapes = null;
                    try
                    {
                        slide = slides.Item(si);
                        shapes = slide.Shapes;
                        int shapeCount = Convert.ToInt32(shapes.Count);

                        for (int shi = 1; shi <= shapeCount; shi++)
                        {
                            dynamic? shape = null;
                            dynamic? actionSettings = null;
                            dynamic? actionSetting = null;
                            dynamic? hyperlink = null;
                            try
                            {
                                shape = shapes.Item(shi);
                                actionSettings = shape.ActionSettings;
                                actionSetting = actionSettings.Item(1); // ppMouseClick = 1
                                int action = Convert.ToInt32(actionSetting.Action);

                                if (action == 7) // ppActionHyperlink
                                {
                                    hyperlink = actionSetting.Hyperlink;
                                    string address = hyperlink.Address?.ToString() ?? "";
                                    string subAddress = hyperlink.SubAddress?.ToString() ?? "";
                                    string screenTip = "";
                                    try { screenTip = hyperlink.ScreenTip?.ToString() ?? ""; } catch { }
                                    string shapeName = shape.Name?.ToString() ?? "";

                                    result.Hyperlinks.Add(new HyperlinkInfo
                                    {
                                        Index = globalIndex++,
                                        Address = address,
                                        SubAddress = subAddress,
                                        ScreenTip = screenTip,
                                        SlideIndex = si,
                                        ShapeName = shapeName
                                    });
                                }
                            }
                            finally
                            {
                                if (hyperlink != null) ComUtilities.Release(ref hyperlink!);
                                if (actionSetting != null) ComUtilities.Release(ref actionSetting!);
                                if (actionSettings != null) ComUtilities.Release(ref actionSettings!);
                                if (shape != null) ComUtilities.Release(ref shape!);
                            }
                        }
                    }
                    finally
                    {
                        if (shapes != null) ComUtilities.Release(ref shapes!);
                        if (slide != null) ComUtilities.Release(ref slide!);
                    }
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref slides!);
            }
        });
    }

    public OperationResult Validate(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = (dynamic)ctx.Presentation;
            dynamic slides = pres.Slides;
            try
            {
                int slideCount = Convert.ToInt32(slides.Count);
                var lines = new List<string>();
                int totalLinks = 0;
                int brokenCount = 0;

                for (int si = 1; si <= slideCount; si++)
                {
                    dynamic? slide = null;
                    dynamic? shapes = null;
                    try
                    {
                        slide = slides.Item(si);
                        shapes = slide.Shapes;
                        int shapeCount = Convert.ToInt32(shapes.Count);

                        for (int shi = 1; shi <= shapeCount; shi++)
                        {
                            dynamic? shape = null;
                            dynamic? actionSettings = null;
                            dynamic? actionSetting = null;
                            dynamic? hyperlink = null;
                            try
                            {
                                shape = shapes.Item(shi);
                                actionSettings = shape.ActionSettings;
                                actionSetting = actionSettings.Item(1); // ppMouseClick = 1
                                int action = Convert.ToInt32(actionSetting.Action);

                                if (action == 7) // ppActionHyperlink
                                {
                                    hyperlink = actionSetting.Hyperlink;
                                    string address = hyperlink.Address?.ToString() ?? "";
                                    string subAddress = hyperlink.SubAddress?.ToString() ?? "";
                                    string shapeName = shape.Name?.ToString() ?? "";
                                    totalLinks++;

                                    string status;
                                    if (string.IsNullOrWhiteSpace(address) && string.IsNullOrWhiteSpace(subAddress))
                                    {
                                        status = "empty";
                                        brokenCount++;
                                    }
                                    else if (!string.IsNullOrWhiteSpace(subAddress) && string.IsNullOrWhiteSpace(address))
                                    {
                                        // Internal slide link (subAddress only, e.g. slide number)
                                        status = "internal";
                                    }
                                    else if (address.StartsWith('#'))
                                    {
                                        status = "internal";
                                    }
                                    else if (address.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                                    {
                                        status = "external";
                                    }
                                    else
                                    {
                                        // Treat as file path
                                        if (System.IO.File.Exists(address))
                                        {
                                            status = "valid";
                                        }
                                        else
                                        {
                                            status = "broken";
                                            brokenCount++;
                                        }
                                    }

                                    string target = !string.IsNullOrWhiteSpace(address) ? address : subAddress;
                                    lines.Add($"Slide {si}, Shape '{shapeName}': {target} [{status}]");
                                }
                            }
                            finally
                            {
                                if (hyperlink != null) ComUtilities.Release(ref hyperlink!);
                                if (actionSetting != null) ComUtilities.Release(ref actionSetting!);
                                if (actionSettings != null) ComUtilities.Release(ref actionSettings!);
                                if (shape != null) ComUtilities.Release(ref shape!);
                            }
                        }
                    }
                    finally
                    {
                        if (shapes != null) ComUtilities.Release(ref shapes!);
                        if (slide != null) ComUtilities.Release(ref slide!);
                    }
                }

                string summary = $"Validated {totalLinks} hyperlink(s): {brokenCount} broken/empty.";
                if (lines.Count > 0)
                    summary += "\n" + string.Join("\n", lines);

                return new OperationResult
                {
                    Success = true,
                    Action = "validate",
                    Message = summary,
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slides!);
            }
        });
    }
}
