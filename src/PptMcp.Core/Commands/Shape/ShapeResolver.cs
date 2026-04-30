using System.Globalization;
using PptMcp.ComInterop;

namespace PptMcp.Core.Commands.Shape;

/// <summary>
/// Resolves a shape on a slide by either its mutable Name (the COM
/// <c>Shape.Name</c> property) or by its stable Id via the <c>id:&lt;N&gt;</c>
/// prefix syntax (matched against the COM <c>Shape.Id</c> property).
/// </summary>
/// <remarks>
/// Centralising the lookup keeps every <see cref="ShapeCommands"/> action
/// consistent: a single place owns the prefix parsing, the COM enumeration,
/// and the descriptive not-found error.
///
/// <para>
/// PowerPoint's <c>Shape.Id</c> is assigned on insertion and is stable for
/// the lifetime of the shape. <c>Shape.Name</c> is user-mutable. Agents
/// that store an Id reference (e.g. from <c>shape(list)</c> or
/// <c>shape(read)</c>) can therefore safely target the same shape across
/// renames, copy/paste, and arbitrary reordering.
/// </para>
/// </remarks>
internal static class ShapeResolver
{
    private const string IdPrefix = "id:";

    /// <summary>
    /// Resolve a shape on the given slide. Returns the COM shape object
    /// — caller is responsible for releasing it via
    /// <see cref="ComUtilities.Release"/>.
    /// </summary>
    /// <param name="slide">The COM slide object that owns the shape.</param>
    /// <param name="shapeNameOrId">
    /// Either the literal <c>Shape.Name</c>, or <c>id:&lt;N&gt;</c> where
    /// <c>N</c> is the integer <c>Shape.Id</c>.
    /// </param>
    /// <exception cref="ArgumentException">
    /// Thrown when <paramref name="shapeNameOrId"/> uses the <c>id:</c>
    /// prefix but the suffix is not a valid integer.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    /// Thrown when an <c>id:&lt;N&gt;</c> reference does not match any
    /// shape on the slide.
    /// </exception>
    public static dynamic Resolve(dynamic slide, string shapeNameOrId)
    {
        ArgumentNullException.ThrowIfNull(shapeNameOrId);

        if (!TryParseId(shapeNameOrId, out int targetId))
        {
            // Existing behavior — let COM throw if the name is unknown.
            return slide.Shapes.Item(shapeNameOrId);
        }

        dynamic? shapes = null;
        try
        {
            shapes = slide.Shapes;
            int count = (int)shapes.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic candidate = shapes.Item(i);
                bool match;
                try
                {
                    match = (int)candidate.Id == targetId;
                }
                catch
                {
                    ComUtilities.Release(ref candidate!);
                    throw;
                }

                if (match)
                {
                    return candidate;
                }

                ComUtilities.Release(ref candidate!);
            }
        }
        finally
        {
            if (shapes != null)
            {
                ComUtilities.Release(ref shapes!);
            }
        }

        throw new InvalidOperationException(
            $"Shape with id '{targetId}' not found on this slide. " +
            $"Use shape(list) to see the current ShapeId values.");
    }

    /// <summary>
    /// Returns <c>true</c> and writes the parsed integer to
    /// <paramref name="id"/> when <paramref name="value"/> matches the
    /// <c>id:&lt;N&gt;</c> shape-reference syntax. Throws
    /// <see cref="ArgumentException"/> when the prefix is present but the
    /// suffix is not a valid integer.
    /// </summary>
    private static bool TryParseId(string value, out int id)
    {
        id = 0;
        if (!value.StartsWith(IdPrefix, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        string suffix = value[IdPrefix.Length..];
        if (!int.TryParse(suffix, NumberStyles.Integer, CultureInfo.InvariantCulture, out id))
        {
            throw new ArgumentException(
                $"Invalid shape reference '{value}'. The 'id:' prefix must be followed " +
                $"by an integer ShapeId (e.g. 'id:42'). Use the literal shape name to " +
                $"reference a shape whose name happens to start with 'id:'.",
                nameof(value));
        }

        return true;
    }
}
