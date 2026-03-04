using PptMcp.ComInterop;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Slide;

/// <summary>
/// Helper methods for reading shape information from PowerPoint COM objects.
/// Shared between SlideCommands and ShapeCommands.
/// </summary>
internal static class ShapeHelpers
{
    /// <summary>
    /// Read shape metadata from a PowerPoint COM Shape object.
    /// </summary>
    internal static ShapeInfo ReadShapeInfo(dynamic shape)
    {
        var info = new ShapeInfo
        {
            ShapeId = (int)shape.Id,
            Name = shape.Name?.ToString() ?? "",
            Left = Convert.ToSingle(shape.Left),
            Top = Convert.ToSingle(shape.Top),
            Width = Convert.ToSingle(shape.Width),
            Height = Convert.ToSingle(shape.Height),
            ZOrderPosition = (int)shape.ZOrderPosition,
        };

        // ShapeType is an MsoShapeType enum, use integer
        int shapeType = Convert.ToInt32(shape.Type);
        info.ShapeType = GetShapeTypeName(shapeType);

        try { info.AlternativeText = shape.AlternativeText?.ToString(); } catch { }

        // HasTextFrame
        try
        {
            info.HasTextFrame = Convert.ToInt32(shape.HasTextFrame) != 0; // msoTrue = -1
            if (info.HasTextFrame)
            {
                try { info.Text = shape.TextFrame.TextRange.Text?.ToString(); } catch { }
            }
        }
        catch { info.HasTextFrame = false; }

        // HasTable
        try { info.HasTable = Convert.ToInt32(shape.HasTable) != 0; } catch { info.HasTable = false; }

        // HasChart
        try { info.HasChart = Convert.ToInt32(shape.HasChart) != 0; } catch { info.HasChart = false; }

        // IsGroup (MsoShapeType.msoGroup = 6)
        info.IsGroup = shapeType == 6;

        // IsPlaceholder (MsoShapeType.msoPlaceholder = 14)
        info.IsPlaceholder = shapeType == 14;
        if (info.IsPlaceholder)
        {
            try { info.PlaceholderType = Convert.ToInt32(shape.PlaceholderFormat.Type); } catch { }
        }

        // Group items
        if (info.IsGroup)
        {
            try
            {
                dynamic groupItems = shape.GroupItems;
                int count = (int)groupItems.Count;
                info.GroupItems = new List<ShapeInfo>(count);
                for (int i = 1; i <= count; i++)
                {
                    dynamic child = groupItems.Item(i);
                    try
                    {
                        info.GroupItems.Add(ReadShapeInfo(child));
                    }
                    finally
                    {
                        ComUtilities.Release(ref child!);
                    }
                }
                ComUtilities.Release(ref groupItems!);
            }
            catch { }
        }

        return info;
    }

    internal static string GetShapeTypeName(int msoShapeType) => msoShapeType switch
    {
        1 => "AutoShape",
        2 => "Callout",
        3 => "Chart",
        4 => "Comment",
        5 => "FreeForm",
        6 => "Group",
        7 => "EmbeddedOLEObject",
        8 => "FormControl",
        9 => "Line",
        10 => "LinkedOLEObject",
        11 => "LinkedPicture",
        12 => "OLEControlObject",
        13 => "Picture",
        14 => "Placeholder",
        15 => "TextEffect",
        16 => "MediaObject",
        17 => "TextBox",
        19 => "Table",
        20 => "Canvas",
        21 => "Diagram",
        22 => "Ink",
        23 => "InkComment",
        24 => "SmartArt",
        25 => "Slicer",
        26 => "WebVideo",
        27 => "ContentApp",
        28 => "Graphic",
        29 => "LinkedGraphic",
        30 => "3DModel",
        31 => "Linked3DModel",
        _ => $"Unknown({msoShapeType})"
    };
}
