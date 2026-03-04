using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Shape;

/// <summary>
/// Shape management: list, read, create, move, resize, delete, z-order.
/// </summary>
[ServiceCategory("shape")]
[McpTool("shape", Title = "Shape Operations", Destructive = true, Category = "shapes")]
public interface IShapeCommands
{
    /// <summary>
    /// List all shapes on a slide.
    /// </summary>
    [ServiceAction("list")]
    ShapeListResult List(IPptBatch batch, int slideIndex);

    /// <summary>
    /// Get detailed info about a specific shape.
    /// </summary>
    [ServiceAction("read")]
    ShapeDetailResult Read(IPptBatch batch, int slideIndex, string shapeName);

    /// <summary>
    /// Add a textbox shape.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="left">Position from left in points</param>
    /// <param name="top">Position from top in points</param>
    /// <param name="width">Width in points</param>
    /// <param name="height">Height in points</param>
    /// <param name="text">Initial text content</param>
    [ServiceAction("add-textbox")]
    OperationResult AddTextbox(IPptBatch batch, int slideIndex, float left, float top, float width, float height, string text);

    /// <summary>
    /// Add a rectangle, ellipse, or other auto-shape.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="autoShapeType">MsoAutoShapeType integer (1=Rectangle, 9=Oval, etc.)</param>
    /// <param name="left">Position from left in points</param>
    /// <param name="top">Position from top in points</param>
    /// <param name="width">Width in points</param>
    /// <param name="height">Height in points</param>
    [ServiceAction("add-shape")]
    OperationResult AddShape(IPptBatch batch, int slideIndex, int autoShapeType, float left, float top, float width, float height);

    /// <summary>
    /// Move and/or resize a shape.
    /// </summary>
    [ServiceAction("move-resize")]
    OperationResult MoveResize(IPptBatch batch, int slideIndex, string shapeName, float? left, float? top, float? width, float? height);

    /// <summary>
    /// Delete a shape by name.
    /// </summary>
    [ServiceAction("delete")]
    OperationResult Delete(IPptBatch batch, int slideIndex, string shapeName);

    /// <summary>
    /// Change the z-order of a shape (bring to front, send to back, etc.).
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the shape</param>
    /// <param name="zOrderCmd">1=BringToFront, 2=SendToBack, 3=BringForward, 4=SendBackward</param>
    [ServiceAction("z-order")]
    OperationResult ZOrder(IPptBatch batch, int slideIndex, string shapeName, int zOrderCmd);

    /// <summary>
    /// Set the fill color of a shape. Use 'none' to remove fill (transparent).
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the shape</param>
    /// <param name="colorHex">Hex color string like #FF0000 for red, or 'none' for no fill</param>
    [ServiceAction("set-fill")]
    OperationResult SetFill(IPptBatch batch, int slideIndex, string shapeName, string colorHex);

    /// <summary>
    /// Set the line/border color and width of a shape. Use 'none' to remove the line.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the shape</param>
    /// <param name="colorHex">Hex color like #000000 or 'none' to remove border</param>
    /// <param name="lineWidth">Line width in points (default 0.75)</param>
    [ServiceAction("set-line")]
    OperationResult SetLine(IPptBatch batch, int slideIndex, string shapeName, string colorHex, float lineWidth);

    /// <summary>
    /// Set the rotation angle of a shape in degrees.
    /// </summary>
    [ServiceAction("set-rotation")]
    OperationResult SetRotation(IPptBatch batch, int slideIndex, string shapeName, float degrees);

    /// <summary>
    /// Group multiple shapes into a single group shape.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeNames">Comma-separated list of shape names to group</param>
    [ServiceAction("group")]
    OperationResult Group(IPptBatch batch, int slideIndex, string shapeNames);

    /// <summary>
    /// Ungroup a group shape into individual shapes.
    /// </summary>
    [ServiceAction("ungroup")]
    OperationResult Ungroup(IPptBatch batch, int slideIndex, string shapeName);

    /// <summary>
    /// Set the alternative text (alt text) of a shape for accessibility.
    /// </summary>
    [ServiceAction("set-alt-text")]
    OperationResult SetAltText(IPptBatch batch, int slideIndex, string shapeName, string altText);

    /// <summary>
    /// Copy a shape to another slide.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based source slide index</param>
    /// <param name="shapeName">Name of the shape to copy</param>
    /// <param name="targetSlideIndex">1-based target slide index</param>
    [ServiceAction("copy-to-slide")]
    OperationResult CopyToSlide(IPptBatch batch, int slideIndex, string shapeName, int targetSlideIndex);

    /// <summary>
    /// Set shadow effect on a shape. Use visible=false to remove shadow.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Shape name</param>
    /// <param name="visible">Show or hide shadow</param>
    /// <param name="offsetX">Shadow offset X in points</param>
    /// <param name="offsetY">Shadow offset Y in points</param>
    [ServiceAction("set-shadow")]
    OperationResult SetShadow(IPptBatch batch, int slideIndex, string shapeName, bool visible, float offsetX, float offsetY);

    /// <summary>
    /// Add a connector line between two shapes.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="connectorType">1=Straight, 2=Elbow, 3=Curve</param>
    /// <param name="startShapeName">Starting shape name</param>
    /// <param name="endShapeName">Ending shape name</param>
    [ServiceAction("add-connector")]
    OperationResult AddConnector(IPptBatch batch, int slideIndex, int connectorType, string startShapeName, string endShapeName);

    /// <summary>
    /// Merge shapes using boolean operations.
    /// mergeType: 1=Union, 2=Combine, 3=Fragment, 4=Intersect, 5=Subtract
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeNames">Comma-separated shape names to merge</param>
    /// <param name="mergeType">1=Union, 2=Combine, 3=Fragment, 4=Intersect, 5=Subtract</param>
    [ServiceAction("merge")]
    OperationResult MergeShapes(IPptBatch batch, int slideIndex, string shapeNames, int mergeType);
    /// <summary>
    /// Duplicate a shape on the same slide.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the shape to duplicate</param>
    [ServiceAction("duplicate")]
    OperationResult Duplicate(IPptBatch batch, int slideIndex, string shapeName);

    /// <summary>
    /// Flip a shape horizontally or vertically.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Shape name</param>
    /// <param name="flipType">0=Horizontal, 1=Vertical</param>
    [ServiceAction("flip")]
    OperationResult Flip(IPptBatch batch, int slideIndex, string shapeName, int flipType);

    /// <summary>
    /// Set TextFrame properties of a shape (margins, word wrap, auto size).
    /// Margins are in points. Pass null to leave a property unchanged.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the shape</param>
    /// <param name="marginLeft">Left margin in points (null = don't change)</param>
    /// <param name="marginRight">Right margin in points (null = don't change)</param>
    /// <param name="marginTop">Top margin in points (null = don't change)</param>
    /// <param name="marginBottom">Bottom margin in points (null = don't change)</param>
    /// <param name="wordWrap">Enable/disable word wrap (null = don't change)</param>
    /// <param name="autoSize">0=None, 1=ShapeToFitText, 2=TextToFitShape (null = don't change)</param>
    [ServiceAction("set-text-frame")]
    OperationResult SetTextFrame(IPptBatch batch, int slideIndex, string shapeName, float? marginLeft, float? marginRight, float? marginTop, float? marginBottom, bool? wordWrap, int? autoSize);

    /// <summary>
    /// Apply a two-color gradient fill to a shape.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the shape</param>
    /// <param name="color1">First gradient color as hex (#RRGGBB)</param>
    /// <param name="color2">Second gradient color as hex (#RRGGBB)</param>
    /// <param name="gradientStyle">1=Horizontal, 2=Vertical, 3=DiagonalUp, 4=DiagonalDown, 5=FromCorner, 6=FromCenter</param>
    [ServiceAction("set-gradient-fill")]
    OperationResult SetGradientFill(IPptBatch batch, int slideIndex, string shapeName, string color1, string color2, int gradientStyle);

    /// <summary>
    /// Set glow effect on a shape. Use radius=0 to remove glow.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the shape</param>
    /// <param name="radius">Glow radius in points (0 = remove glow)</param>
    /// <param name="colorHex">Glow color as hex (#RRGGBB)</param>
    [ServiceAction("set-glow")]
    OperationResult SetGlow(IPptBatch batch, int slideIndex, string shapeName, float radius, string colorHex);

    /// <summary>
    /// Set reflection effect on a shape. Use reflectionType=0 to remove reflection.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the shape</param>
    /// <param name="reflectionType">0=None, 1-9=msoReflectionType1 through msoReflectionType9</param>
    [ServiceAction("set-reflection")]
    OperationResult SetReflection(IPptBatch batch, int slideIndex, string shapeName, int reflectionType);

    /// <summary>
    /// Set the opacity (transparency) of a shape's fill.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the shape</param>
    /// <param name="opacity">Opacity value from 0.0 (fully transparent) to 1.0 (fully opaque)</param>
    [ServiceAction("set-opacity")]
    OperationResult SetOpacity(IPptBatch batch, int slideIndex, string shapeName, float opacity);

    /// <summary>
    /// Read the fill properties of a shape: fill type, color (if solid), and transparency.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the shape</param>
    [ServiceAction("read-fill")]
    OperationResult ReadFill(IPptBatch batch, int slideIndex, string shapeName);

    /// <summary>
    /// Read the line/border properties of a shape: visible, color, weight.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the shape</param>
    [ServiceAction("read-line")]
    OperationResult ReadLine(IPptBatch batch, int slideIndex, string shapeName);

    /// <summary>
    /// Find all shapes on a slide that match a given MsoShapeType.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeType">MsoShapeType integer (1=AutoShape, 6=Group, 13=Picture, 14=Placeholder, 17=TextBox, etc.)</param>
    [ServiceAction("find-by-type")]
    OperationResult FindByType(IPptBatch batch, int slideIndex, int shapeType);

    /// <summary>
    /// Copy all formatting from one shape to another using Format Painter (PickUp/Apply).
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="sourceShapeName">Name of the shape to copy formatting from</param>
    /// <param name="targetShapeName">Name of the shape to apply formatting to</param>
    [ServiceAction("copy-formatting")]
    OperationResult CopyFormatting(IPptBatch batch, int slideIndex, string sourceShapeName, string targetShapeName);
}
