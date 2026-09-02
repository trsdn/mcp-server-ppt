using System.ComponentModel;
using System.Text.Json;
using PptMcp.CLI.Infrastructure;
using PptMcp.Core.Commands.File;
using PptMcp.Service;
using Spectre.Console;
using Spectre.Console.Cli;

namespace PptMcp.CLI.Commands;

// ============================================================================
// SESSION COMMANDS
// ============================================================================

internal sealed class SessionCreateCommand : AsyncCommand<SessionCreateCommand.Settings>
{
    public override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.FilePath))
        {
            AnsiConsole.MarkupLine("[red]File path is required.[/]");
            return 1;
        }

        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "session.create",
            Args = JsonSerializer.Serialize(new { filePath = settings.FilePath, timeoutSeconds = settings.TimeoutSeconds }, ServiceProtocol.JsonOptions)
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(response.Result);
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, ServiceProtocol.JsonOptions));
            return 1;
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<FILE>")]
        [Description("Path to the new PowerPoint file to create")]
        public string FilePath { get; init; } = string.Empty;

        [CommandOption("--timeout <SECONDS>")]
        [Description("Session timeout in seconds")]
        public int? TimeoutSeconds { get; init; }
    }
}

internal sealed class SessionOpenCommand : AsyncCommand<SessionOpenCommand.Settings>
{
    public override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.FilePath))
        {
            AnsiConsole.MarkupLine("[red]File path is required.[/]");
            return 1;
        }

        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = settings.FilePath, timeoutSeconds = settings.TimeoutSeconds }, ServiceProtocol.JsonOptions)
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(response.Result);
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, ServiceProtocol.JsonOptions));
            return 1;
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<FILE>")]
        [Description("Path to the PowerPoint file to open")]
        public string FilePath { get; init; } = string.Empty;

        [CommandOption("--timeout <SECONDS>")]
        [Description("Session timeout in seconds")]
        public int? TimeoutSeconds { get; init; }
    }
}

internal sealed class SessionCloseCommand : AsyncCommand<SessionCloseCommand.Settings>
{
    public override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.SessionId))
        {
            AnsiConsole.MarkupLine("[red]Session ID is required.[/]");
            return 1;
        }

        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = settings.SessionId,
            Args = JsonSerializer.Serialize(new { save = settings.Save }, ServiceProtocol.JsonOptions)
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = settings.Save ? "Session closed and saved." : "Session closed." }, ServiceProtocol.JsonOptions));
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, ServiceProtocol.JsonOptions));
            return 1;
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandOption("-s|--session <SESSION>")]
        [Description("Session ID to close")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--save")]
        [Description("Save changes before closing")]
        public bool Save { get; init; }
    }
}

internal sealed class SessionListCommand : AsyncCommand
{
    public override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        var pipeName = DaemonAutoStart.GetPipeName();
        using var client = new ServiceClient(pipeName, connectTimeout: TimeSpan.FromSeconds(2));

        try
        {
            var response = await client.SendAsync(new ServiceRequest { Command = "session.list" }, cancellationToken);
            if (response.Success)
            {
                Console.WriteLine(response.Result);
                return 0;
            }
            else
            {
                Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, ServiceProtocol.JsonOptions));
                return 1;
            }
        }
        catch (Exception)
        {
            // Daemon not running — no sessions
            Console.WriteLine(JsonSerializer.Serialize(new { sessions = Array.Empty<object>() }, ServiceProtocol.JsonOptions));
            return 0;
        }
    }
}

internal sealed class SessionSaveCommand : AsyncCommand<SessionSaveCommand.Settings>
{
    public override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.SessionId))
        {
            AnsiConsole.MarkupLine("[red]Session ID is required.[/]");
            return 1;
        }

        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "session.save",
            SessionId = settings.SessionId
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = "Session saved." }, ServiceProtocol.JsonOptions));
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, ServiceProtocol.JsonOptions));
            return 1;
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandOption("-s|--session <SESSION>")]
        [Description("Session ID to save")]
        public string SessionId { get; init; } = string.Empty;
    }
}





// ============================================================================
// SESSION TEST
//
// Inspects a file without opening it. Present on the MCP 'file' tool since its
// introduction; absent from the CLI until issue #131, which is exactly the drift
// that can happen when a surface is hand-written on both sides rather than
// generated from a shared interface.
//
// Deliberately does not go through the daemon. FileCommands.Test is pure file
// inspection - it never launches PowerPoint - so routing it through a pipe would
// start a daemon to answer a question about the file system, and would make the
// preflight check fail in exactly the situations where a preflight is most useful.
// ============================================================================

internal sealed class SessionTestCommand : Command<SessionTestCommand.Settings>
{
    public override int Execute(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.FilePath))
        {
            AnsiConsole.MarkupLine("[red]File path is required.[/]");
            return 1;
        }

        var info = new FileCommands().Test(settings.FilePath);

        Console.WriteLine(JsonSerializer.Serialize(new
        {
            success = info.Success,
            exists = info.Exists,
            filePath = info.FilePath,
            fileName = info.FileName,
            fileSizeBytes = info.FileSizeBytes,
            isReadOnly = info.IsReadOnly,
            isMacroEnabled = info.IsMacroEnabled
        }, ServiceProtocol.JsonOptions));

        return info.Success ? 0 : 1;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<FILE>")]
        [Description("Absolute path to the .pptx/.pptm file to inspect")]
        public string FilePath { get; init; } = string.Empty;
    }
}
