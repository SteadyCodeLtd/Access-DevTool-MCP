using System.Diagnostics;
using System.Text.Json;
using AccessAceMcp;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AccessAceMcpTests;

// Tests for the named-pipe worker protocol in Access-ACE-Agent.exe.
// These tests start the real worker process but do NOT require Microsoft Access
// to be installed — they only exercise code paths that don't reach COM.

[TestClass]
[TestCategory("WorkerProtocol")]
[DoNotParallelize]
public class WorkerPipeProtocolTests
{
    // ── Worker lifecycle helpers ──────────────────────────────────────────────

    private static string FindWorkerExe()
    {
        var dir = AppContext.BaseDirectory;
        var candidates = new[]
        {
            // Placed next to test exe via post-build (preferred)
            Path.Combine(dir, "Access-ACE-Agent.exe"),
            // Debug build relative to test output (net10.0/Debug)
            Path.GetFullPath(Path.Combine(dir, "..", "..", "..", "..", "Access_ACE_Agent", "bin", "Debug", "Access-ACE-Agent.exe")),
            Path.GetFullPath(Path.Combine(dir, "..", "..", "..", "..", "Access_ACE_Agent", "bin", "Release", "Access-ACE-Agent.exe")),
        };

        var found = candidates.FirstOrDefault(File.Exists);
        if (found == null)
            Assert.Inconclusive(
                "Access-ACE-Agent.exe not found. Build the Access_ACE_Agent project first. " +
                $"Searched: {string.Join(", ", candidates)}");

        return found!;
    }

    private sealed class WorkerFixture : IAsyncDisposable
    {
        public PipeChannel Channel { get; }
        private readonly Process _process;

        private WorkerFixture(Process p, PipeChannel c) { _process = p; Channel = c; }

        public static WorkerFixture Start(string workerExe, string? dbPath = null)
        {
            var pipeName = $"proto-test-{Guid.NewGuid():N}";
            var arguments = $"--pipe {pipeName}";
            if (dbPath != null) arguments += $" \"{dbPath}\"";

            var process = Process.Start(new ProcessStartInfo(workerExe)
            {
                Arguments            = arguments,
                UseShellExecute      = false,
                RedirectStandardError = true,
                CreateNoWindow       = true,
            }) ?? throw new InvalidOperationException("Failed to start worker process");

            process.ErrorDataReceived += (_, e) =>
            {
                if (e.Data is { } line) Console.Error.WriteLine($"[worker-stderr] {line}");
            };
            process.BeginErrorReadLine();

            var channel = new PipeChannel(pipeName, connectTimeoutMs: 15_000);
            return new WorkerFixture(process, channel);
        }

        public async ValueTask DisposeAsync()
        {
            Channel.Dispose();
            try { if (!_process.HasExited) _process.Kill(entireProcessTree: true); }
            catch { /* best-effort */ }
            await _process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5)).ContinueWith(_ => { });
            _process.Dispose();
        }
    }

    private static JsonElement ParseResult(string json)
    {
        var doc = JsonDocument.Parse(json);
        return doc.RootElement.Clone(); // clone so we can use after doc disposal
    }

    // ── Protocol tests (no Access needed) ────────────────────────────────────

    [TestMethod]
    public async Task IsConnected_WhenNotConnected_ReturnsFalseWithSuccess()
    {
        await using var w = WorkerFixture.Start(FindWorkerExe());

        var result = ParseResult(await w.Channel.CallAsync("is_connected"));

        Assert.IsTrue(result.GetProperty("success").GetBoolean());
        Assert.IsFalse(result.GetProperty("connected").GetBoolean());
    }

    [TestMethod]
    public async Task ConnectAccess_NonExistentPath_ReturnsErrorResult()
    {
        await using var w = WorkerFixture.Start(FindWorkerExe());

        var result = ParseResult(await w.Channel.CallAsync(
            "connect_access", new { database_path = @"C:\does\not\exist\fake.accdb" }));

        Assert.IsFalse(result.GetProperty("success").GetBoolean(),
            "Connecting to a missing file should return success:false");
        Assert.IsTrue(result.TryGetProperty("error", out _),
            "Error result must include an 'error' field");
    }

    [TestMethod]
    public async Task UnknownTool_ReturnsErrorResult()
    {
        await using var w = WorkerFixture.Start(FindWorkerExe());

        var result = ParseResult(await w.Channel.CallAsync("no_such_tool_xyzzy", new { }));

        Assert.IsFalse(result.GetProperty("success").GetBoolean());
        var error = result.GetProperty("error").GetString() ?? "";
        StringAssert.Contains(error, "no_such_tool_xyzzy");
    }

    [TestMethod]
    public async Task RunCommand_IntegerCommandId_NotAParseError()
    {
        // Verifies ParseIntArg handles JSON number (not string) for command_id.
        // The command -1 is invalid so Access rejects it, but the error must be
        // a domain error ("command") not a type-parsing error ("must be an integer").
        await using var w = WorkerFixture.Start(FindWorkerExe());

        var result = ParseResult(await w.Channel.CallAsync("run_command", new { command_id = -1 }));

        Assert.IsFalse(result.GetProperty("success").GetBoolean());
        var error = result.GetProperty("error").GetString() ?? "";
        Assert.IsFalse(error.Contains("must be an integer"),
            $"Got a type-parse error instead of domain error: {error}");
    }

    [TestMethod]
    public async Task ExportObjectToText_IntegerObjectType_NotAParseError()
    {
        // Verifies ParseIntArg handles JSON number for object_type.
        // Tool will fail (not connected) but must not fail with a type-parse error.
        await using var w = WorkerFixture.Start(FindWorkerExe());

        var result = ParseResult(await w.Channel.CallAsync(
            "export_object_to_text", new { object_type = 2, object_name = "FakeForm" }));

        Assert.IsFalse(result.GetProperty("success").GetBoolean());
        var error = result.GetProperty("error").GetString() ?? "";
        Assert.IsFalse(error.Contains("must be an integer"),
            $"Got a type-parse error instead of domain error: {error}");
    }

    [TestMethod]
    public async Task MultipleSequentialCalls_AllSucceed()
    {
        await using var w = WorkerFixture.Start(FindWorkerExe());

        for (int i = 0; i < 5; i++)
        {
            var result = ParseResult(await w.Channel.CallAsync("is_connected"));
            Assert.IsTrue(result.GetProperty("success").GetBoolean(),
                $"Call {i} failed unexpectedly");
        }
    }

    [TestMethod]
    public async Task DisconnectAccess_WhenNotConnected_DoesNotCrash()
    {
        // disconnect when never connected should return a graceful result (success or error),
        // not crash the worker.
        await using var w = WorkerFixture.Start(FindWorkerExe());

        var result = ParseResult(await w.Channel.CallAsync("disconnect_access"));

        // After disconnect attempt, worker must still respond to further calls
        var check = ParseResult(await w.Channel.CallAsync("is_connected"));
        Assert.IsTrue(check.GetProperty("success").GetBoolean(),
            "Worker must still be alive after disconnect-when-not-connected");
    }

    [TestMethod]
    public async Task GetForms_WhenNotConnected_ReturnsErrorNotCrash()
    {
        await using var w = WorkerFixture.Start(FindWorkerExe());

        var result = ParseResult(await w.Channel.CallAsync("get_forms"));

        Assert.IsFalse(result.GetProperty("success").GetBoolean());
        Assert.IsTrue(result.TryGetProperty("error", out _));
    }

    [TestMethod]
    public async Task EmptyParamsObject_HandledSameAsNoParams()
    {
        // Sending params:{} vs no params key should both work for is_connected
        await using var w = WorkerFixture.Start(FindWorkerExe());

        var r1 = ParseResult(await w.Channel.CallAsync("is_connected"));
        var r2 = ParseResult(await w.Channel.CallAsync("is_connected", new { }));

        Assert.AreEqual(
            r1.GetProperty("success").GetBoolean(),
            r2.GetProperty("success").GetBoolean());
        Assert.AreEqual(
            r1.GetProperty("connected").GetBoolean(),
            r2.GetProperty("connected").GetBoolean());
    }
}
