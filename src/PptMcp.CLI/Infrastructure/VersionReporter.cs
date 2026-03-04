using System.Reflection;
using Spectre.Console;

namespace PptMcp.CLI.Infrastructure;

internal static class VersionReporter
{
    public static void WriteVersion()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var version = assembly.GetName().Version;
        var informational = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
                           ?? version?.ToString()
                           ?? "unknown";

        AnsiConsole.MarkupLine($"[bold cyan]PptMcp.CLI[/] [green]v{informational}[/]");
        AnsiConsole.MarkupLine($"[dim]Runtime:[/] {System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription}");
        AnsiConsole.MarkupLine($"[dim]Platform:[/] {System.Runtime.InteropServices.RuntimeInformation.OSDescription}");
        AnsiConsole.MarkupLine("[bold]Repository:[/] https://github.com/trsdn/mcp-server-ppt");
        AnsiConsole.MarkupLine("[bold]License:[/] MIT");
    }
}


