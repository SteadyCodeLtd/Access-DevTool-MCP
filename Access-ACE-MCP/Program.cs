using System.Diagnostics;
using AccessAceMcp;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using ModelContextProtocol.Server;

// Locate the .NET 4.8 COM worker next to this executable
var agentExe = Path.Combine(AppContext.BaseDirectory, "Access-ACE-Agent.exe");
if (!File.Exists(agentExe))
    throw new FileNotFoundException(
        $"Access-ACE-Agent.exe not found at '{agentExe}'. " +
        "Build the Access-ACE-Agent project and copy it alongside Access-ACE-MCP.exe.");

var pipeName = $"access-ace-{Environment.ProcessId}";

// Build worker argument list: --pipe <name> [optional db path from our own args]
var workerArgs = $"--pipe {pipeName}";
if (args.Length > 0 && File.Exists(args[0]))
    workerArgs += $" \"{args[0]}\"";

using var workerProcess = new Process
{
    StartInfo = new ProcessStartInfo(agentExe)
    {
        Arguments        = workerArgs,
        UseShellExecute  = false,
        RedirectStandardError = true,
        CreateNoWindow   = true,
    }
};
workerProcess.ErrorDataReceived += (_, e) =>
{
    if (e.Data is { } line) Console.Error.WriteLine($"[worker] {line}");
};
workerProcess.Start();
workerProcess.BeginErrorReadLine();

// Connect to the worker pipe (blocks until worker starts its server, up to 30 s)
var channel = new PipeChannel(pipeName, connectTimeoutMs: 30_000);

var builder = Host.CreateEmptyApplicationBuilder(new HostApplicationBuilderSettings { Args = args });
#pragma warning disable CA1416 // AccessTools is Windows-only by design — this server only runs on Windows
builder.Services
    .AddSingleton(channel)
    .AddMcpServer()
    .WithStdioServerTransport()
    .WithTools<AccessTools>();
#pragma warning restore CA1416

var host = builder.Build();

// Clean up worker when the host shuts down
host.Services
    .GetRequiredService<IHostApplicationLifetime>()
    .ApplicationStopping.Register(() =>
    {
        channel.Dispose();
        try
        {
            if (!workerProcess.HasExited)
                workerProcess.Kill(entireProcessTree: true);
        }
        catch { /* best-effort */ }
    });

await host.RunAsync();
