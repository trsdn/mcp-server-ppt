using System.IO.Pipelines;
using System.Reflection;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.WorkerService;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using PptMcp.McpServer.Telemetry;

namespace PptMcp.McpServer;

/// <summary>
/// PptMcp Model Context Protocol (MCP) Server.
/// Provides resource-based tools for AI assistants to automate PowerPoint operations.
/// </summary>
public class Program
{
    // Test transport configuration - set by tests before calling Main()
    // These are intentionally static for test injection. Thread-safety is not required
    // because tests run sequentially and call ResetTestTransport() after each test.
    private static Pipe? _testInputPipe;
    private static Pipe? _testOutputPipe;

    /// <summary>
    /// Configures the server to use in-memory pipe transport for testing.
    /// Call this before RunAsync() to enable test mode.
    /// </summary>
    /// <param name="inputPipe">Pipe for reading client requests (client writes, server reads)</param>
    /// <param name="outputPipe">Pipe for writing server responses (server writes, client reads)</param>
    public static void ConfigureTestTransport(Pipe inputPipe, Pipe outputPipe)
    {
        _testInputPipe = inputPipe;
        _testOutputPipe = outputPipe;
    }

    /// <summary>
    /// Resets test transport configuration (call after test completes).
    /// </summary>
    public static void ResetTestTransport()
    {
        _testInputPipe = null;
        _testOutputPipe = null;
    }

    public static async Task<int> Main(string[] args)
    {
        // Register assembly resolver for office.dll (Microsoft.Office.Core), which is a
        // .NET Framework GAC assembly that .NET Core cannot find via standard probing.
        // office.dll is copied to our output directory by Directory.Build.targets.
        RegisterOfficeAssemblyResolver();

        // Handle --help and --version flags for easy verification
        if (args.Length > 0)
        {
            var arg = args[0].ToLowerInvariant();
            if (arg is "-h" or "--help" or "-?" or "/?" or "/h")
            {
                ShowHelp();
                return 0;
            }
            if (arg is "-v" or "--version")
            {
                await ShowVersionAsync();
                return 0;
            }
        }

        // Register global exception handlers for unhandled exceptions (telemetry)
        RegisterGlobalExceptionHandlers();

        var builder = Host.CreateApplicationBuilder(args);

        // Disable FileSystemWatcher for config file reload.
        // Host.CreateApplicationBuilder() enables reloadOnChange:true by default, creating a
        // FileSystemWatcher for appsettings.json. Under file I/O storms (PowerPoint temp files, lock
        // files), this watcher fires ParseEventBufferAndNotifyForEach in a tight loop on the
        // threadpool, consuming ~85% CPU. Since MCP server config never changes at runtime,
        // disable reload entirely to eliminate the watcher.
        // Re-add JSON, environment variables, and CLI args — minus the file watchers.
        builder.Configuration.Sources.Clear();
        builder.Configuration
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: false)
            .AddJsonFile($"appsettings.{builder.Environment.EnvironmentName}.json", optional: true, reloadOnChange: false)
            .AddEnvironmentVariables()
            .AddCommandLine(args);

        // For stdio transport: Clear console logging to avoid polluting stderr with info messages.
        // The MCP client interprets stderr output as errors/warnings, so we only log Warning+
        // to stderr for debugging purposes. The MCP SDK handles protocol-level logging.
        builder.Logging.ClearProviders();
        builder.Logging.AddConsole(consoleLogOptions =>
        {
            // Only log Warning and above to stderr - Info/Debug would appear as errors in MCP clients
            consoleLogOptions.LogToStandardErrorThreshold = LogLevel.Warning;
        });
        builder.Logging.SetMinimumLevel(LogLevel.Warning);

        // Configure Application Insights
        ConfigureTelemetry(builder);

        // Configure MCP Server - use test transport if configured, otherwise stdio
        var mcpBuilder = builder.Services
            .AddMcpServer(options =>
            {
                options.ServerInfo = new()
                {
                    Name = "ppt-mcp",
                    Version = typeof(Program).Assembly.GetName().Version?.ToString() ?? "1.0.0"
                };

                // Server-wide instructions for LLMs - helps with tool selection and workflow understanding
                options.ServerInstructions = """
                    PptMcp automates Microsoft PowerPoint via COM interop.

                    CRITICAL: File must be CLOSED in PowerPoint desktop app (COM requires exclusive access).

                    SESSION LIFECYCLE:
                    1. file(action:'open') → returns session_id
                    2. Use session_id with ALL subsequent tools
                    3. file(action:'close', save:true/false) → ONLY when completely done
                    """;
            })
            .WithToolsFromAssembly()
            .WithPromptsFromAssembly(); // Auto-discover prompts marked with [McpServerPromptType]

        if (_testInputPipe != null && _testOutputPipe != null)
        {
            // Test mode: use in-memory pipe transport
            mcpBuilder.WithStreamServerTransport(
                _testInputPipe.Reader.AsStream(),
                _testOutputPipe.Writer.AsStream());
        }
        else
        {
            // Production mode: use stdio transport
            mcpBuilder.WithStdioServerTransport();
        }

        var host = builder.Build();

        // Initialize telemetry client for static access
        InitializeTelemetryClient(host.Services);

        // Note: Update checks are handled by PptMcp Service (shown via Windows notification)
        // to avoid duplicate notifications when running in unified package mode

        try
        {
            await host.RunAsync();
            return 0;
        }
        catch (OperationCanceledException)
        {
            // Graceful shutdown via cancellation (e.g., Ctrl+C, SIGTERM)
            // This is expected behavior, not an error
            return 0;
        }
#pragma warning disable CA1031 // Catch general exception - this is a top-level handler that must not crash
        catch (Exception ex)
        {
            // Track MCP SDK/transport errors (protocol errors, serialization errors, etc.)
            PptMcpTelemetry.TrackUnhandledException(ex, "McpServer.RunAsync");
            PptMcpTelemetry.Flush(); // Ensure telemetry is sent before exit

            // Return exit code 1 for fatal errors (FR-024, SC-015a)
            // Do NOT re-throw - deterministic exit code is more important for callers
            return 1;
        }
#pragma warning restore CA1031
        finally
        {
            // CRITICAL: Auto-save all sessions and clean up PowerPoint processes on shutdown.
            // Without this, MCP client disconnect or process exit silently discards all unsaved work.
            ServiceBridge.ServiceBridge.Dispose();
        }
    }

    /// <summary>
    /// Initializes the static TelemetryClient from DI container.
    /// </summary>
    private static void InitializeTelemetryClient(IServiceProvider services)
    {
        // Resolve TelemetryClient from DI and store for static access
        // Worker Service SDK manages the TelemetryClient lifecycle including flush on shutdown
        var telemetryClient = services.GetService<TelemetryClient>();
        if (telemetryClient != null)
        {
            PptMcpTelemetry.SetTelemetryClient(telemetryClient);
        }
    }

    /// <summary>
    /// Configures Application Insights Worker Service SDK for telemetry.
    /// Uses AddApplicationInsightsTelemetryWorkerService() for proper host integration.
    /// Enables Users/Sessions/Funnels/User Flows analytics in Azure Portal.
    /// </summary>
    private static void ConfigureTelemetry(HostApplicationBuilder builder)
    {
        var connectionString = PptMcpTelemetry.GetConnectionString();
        if (string.IsNullOrEmpty(connectionString))
        {
            return; // No connection string available (local dev build)
        }

        // Configure Application Insights Worker Service SDK
        // This provides:
        // - Proper DI integration with IHostApplicationLifetime
        // - Automatic dependency tracking
        // - Automatic performance counter collection (where available)
        // - Proper telemetry channel with ServerTelemetryChannel (retries, local storage)
        // - Automatic flush on host shutdown
        var aiOptions = new ApplicationInsightsServiceOptions
        {
            // Set connection string if available
            ConnectionString = connectionString,

            // Disable features not needed for MCP server (reduces overhead)
            EnableHeartbeat = true,  // Useful for monitoring server health
            EnableAdaptiveSampling = true,  // Helps manage telemetry volume
            EnableQuickPulseMetricStream = false,  // Live Metrics not needed for CLI tool
            EnablePerformanceCounterCollectionModule = false,  // Perf counters not useful for short-lived CLI
            EnableEventCounterCollectionModule = false,  // Event counters not needed

            // Disable dependency tracking for HTTP calls
            EnableDependencyTrackingTelemetryModule = false,
        };

        builder.Services.AddApplicationInsightsTelemetryWorkerService(aiOptions);

        // Add custom telemetry initializer for User.Id and Session.Id
        // This enables the Users and Sessions blades in Azure Portal
        builder.Services.AddSingleton<Microsoft.ApplicationInsights.Extensibility.ITelemetryInitializer, PptMcpTelemetryInitializer>();
    }

    /// <summary>
    /// Registers global exception handlers to capture unhandled exceptions.
    /// </summary>
    private static void RegisterOfficeAssemblyResolver()
    {
        AppDomain.CurrentDomain.AssemblyResolve += (_, args) =>
        {
            var name = new AssemblyName(args.Name);
            if (!string.Equals(name.Name, "office", StringComparison.OrdinalIgnoreCase))
                return null;

            return ResolveOfficeDll();
        };
    }

    /// <summary>
    /// Resolves office.dll (Microsoft.Office.Core) from multiple locations.
    /// office.dll is a .NET Framework GAC assembly that .NET Core cannot find automatically.
    /// It is present when Microsoft Office is installed, but not in the .NET Core probing paths.
    /// Search order:
    ///   1. AppContext.BaseDirectory (copied by Directory.Build.targets in local dev builds)
    ///   2. .NET Framework GAC - v16 then v15 (v15 is accepted by the CLR for v16 requests)
    ///   3. Office installation directory (click-to-run Office 365 doesn't register in GAC)
    /// </summary>
    private static Assembly? ResolveOfficeDll()
    {
        // 1. Local build output (Directory.Build.targets copies office.dll here in dev builds)
        var localPath = Path.Combine(AppContext.BaseDirectory, "office.dll");
        if (File.Exists(localPath))
            return Assembly.LoadFrom(localPath);

        // 2. .NET Framework GAC — v16 preferred, v15 accepted (CLR honours AssemblyResolve return regardless of version)
        string[] gacPaths =
        [
            @"C:\Windows\assembly\GAC_MSIL\office\16.0.0.0__71e9bce111e9429c\OFFICE.DLL",
            @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL",
        ];
        foreach (var gacPath in gacPaths)
        {
            if (File.Exists(gacPath))
                return Assembly.LoadFrom(gacPath);
        }

        // 3. Office 365 click-to-run installation directories (Office registers its own copy)
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
        string[] officeDirs =
        [
            Path.Combine(programFiles, @"Microsoft Office\root\Office16"),
            Path.Combine(programFilesX86, @"Microsoft Office\root\Office16"),
        ];
        foreach (var dir in officeDirs)
        {
            var officePath = Path.Combine(dir, "OFFICE.dll");
            if (File.Exists(officePath))
                return Assembly.LoadFrom(officePath);
        }

        return null;
    }

    private static void RegisterGlobalExceptionHandlers()
    {
        // Handle exceptions that escape all catch blocks
        AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
        {
            if (e.ExceptionObject is Exception ex)
            {
                PptMcpTelemetry.TrackUnhandledException(ex, "AppDomain.UnhandledException");
            }
        };

        // Handle unobserved task exceptions
        TaskScheduler.UnobservedTaskException += (sender, e) =>
        {
            PptMcpTelemetry.TrackUnhandledException(e.Exception, "TaskScheduler.UnobservedTaskException");
            // Don't observe it - let the runtime handle it
        };
    }

    /// <summary>
    /// Shows help information.
    /// </summary>
    private static void ShowHelp()
    {
        var version = typeof(Program).Assembly.GetName().Version?.ToString() ?? "1.0.0";
        Console.WriteLine($"""
            PowerPoint MCP Server v{version}

            An MCP (Model Context Protocol) server for Microsoft PowerPoint automation.

            Usage:
              PptMcp.McpServer.exe [options]

            Options:
              -h, --help      Show this help message
              -v, --version   Show version information

            Without options, starts the MCP server in stdio mode.

            Requirements:
              - Windows x64
              - Microsoft PowerPoint 2016 or later (desktop version)
            """);
    }

    /// <summary>
    /// Shows version information and checks for updates.
    /// </summary>
    private static async Task ShowVersionAsync()
    {
        var currentVersion = Infrastructure.McpServerVersionChecker.GetCurrentVersion();
        Console.WriteLine($"PowerPoint MCP Server v{currentVersion}");

        // Check for updates (non-blocking, 5-second timeout)
        var latestVersion = await Infrastructure.McpServerVersionChecker.CheckForUpdateAsync();
        if (latestVersion != null)
        {
            Console.WriteLine();
            Console.WriteLine($"Update available: {currentVersion} -> {latestVersion}");
            Console.WriteLine("Run: dotnet tool update --global PptMcp.McpServer");
            Console.WriteLine("Release notes: https://github.com/trsdn/mcp-server-ppt/releases/latest");
        }
    }
}



