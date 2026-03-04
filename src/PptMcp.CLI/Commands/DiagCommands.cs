using System.ComponentModel;
using System.Text.Json;
using PptMcp.CLI.Infrastructure;
using PptMcp.Service;
using Spectre.Console.Cli;

namespace PptMcp.CLI.Commands;

// ============================================================================
// DIAG COMMANDS - Diagnostic/infrastructure commands (no PowerPoint required)
// ============================================================================

internal sealed class DiagPingCommand : AsyncCommand
{
    public override async Task<int> ExecuteAsync(CommandContext context, CancellationToken cancellationToken)
    {
        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest { Command = "diag.ping" }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(response.Result);
            return 0;
        }

        Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, ServiceProtocol.JsonOptions));
        return 1;
    }
}

internal sealed class DiagEchoCommand : AsyncCommand<DiagEchoCommand.Settings>
{
    public override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.Message))
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = "Parameter 'message' is required for echo" }, ServiceProtocol.JsonOptions));
            return 1;
        }

        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var args = new Dictionary<string, object?> { ["message"] = settings.Message };
        if (!string.IsNullOrEmpty(settings.Tag))
            args["tag"] = settings.Tag;

        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "diag.echo",
            Args = JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions)
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(response.Result);
            return 0;
        }

        Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, ServiceProtocol.JsonOptions));
        return 1;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandOption("--message <MESSAGE>")]
        [Description("Message to echo back")]
        public string? Message { get; init; }

        [CommandOption("--tag <TAG>")]
        [Description("Optional tag to include")]
        public string? Tag { get; init; }
    }
}

internal sealed class DiagValidateParamsCommand : AsyncCommand<DiagValidateParamsCommand.Settings>
{
    public override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.Name))
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = "Parameter 'name' is required for validate-params" }, ServiceProtocol.JsonOptions));
            return 1;
        }

        using var client = await DaemonAutoStart.EnsureAndConnectAsync(cancellationToken);
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = "diag.validate-params",
            Args = JsonSerializer.Serialize(new Dictionary<string, object?>
            {
                ["name"] = settings.Name,
                ["count"] = settings.Count,
                ["label"] = settings.Label,
                ["verbose"] = settings.Verbose
            }, ServiceProtocol.JsonOptions)
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(response.Result);
            return 0;
        }

        Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, ServiceProtocol.JsonOptions));
        return 1;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandOption("--name <NAME>")]
        [Description("Name parameter (required)")]
        public string? Name { get; init; }

        [CommandOption("--count <COUNT>")]
        [Description("Count parameter")]
        [DefaultValue(0)]
        public int Count { get; init; }

        [CommandOption("--label <LABEL>")]
        [Description("Optional label")]
        public string? Label { get; init; }

        [CommandOption("--verbose")]
        [Description("Verbose flag")]
        [DefaultValue(false)]
        public bool Verbose { get; init; }
    }
}
