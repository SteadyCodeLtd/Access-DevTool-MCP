using System.IO.Pipes;
using System.Text;
using System.Text.Json;
using AccessAceMcp;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AccessAceMcpTests;

// Tests for PipeChannel using a lightweight in-process fake pipe server.
// No Access installation or worker process needed.

[TestClass]
[TestCategory("PipeChannel")]
public class PipeChannelTests
{
    // ── Fake server helpers ───────────────────────────────────────────────────

    /// <summary>
    /// Creates a connected server/client pair for a test and returns the
    /// server-side reader and writer so the test can simulate the worker.
    /// </summary>
    private static async Task<(PipeChannel client, StreamReader serverReader, StreamWriter serverWriter, NamedPipeServerStream serverPipe)>
        CreatePairAsync(string pipeName)
    {
        var serverPipe = new NamedPipeServerStream(
            pipeName, PipeDirection.InOut, maxNumberOfServerInstances: 1,
            PipeTransmissionMode.Byte, PipeOptions.Asynchronous);

        // Start accept before creating client (client blocks until server listens)
        var acceptTask = serverPipe.WaitForConnectionAsync();
        var client = new PipeChannel(pipeName, connectTimeoutMs: 5_000);
        await acceptTask;

        var serverReader = new StreamReader(serverPipe, new UTF8Encoding(false), false, 65536, leaveOpen: true);
        var serverWriter = new StreamWriter(serverPipe, new UTF8Encoding(false), 65536, leaveOpen: true)
            { AutoFlush = true };

        return (client, serverReader, serverWriter, serverPipe);
    }

    private static string UniquePipe() => $"pipetest-{Guid.NewGuid():N}";

    // ── Request serialization ─────────────────────────────────────────────────

    [TestMethod]
    public async Task CallAsync_NoParams_SendsOnlyToolKey()
    {
        var (client, sr, sw, pipe) = await CreatePairAsync(UniquePipe());
        try
        {
            var callTask = client.CallAsync("is_connected");

            var line = await sr.ReadLineAsync();
            await sw.WriteLineAsync("{\"success\":true,\"connected\":false}");
            await callTask;

            Assert.IsNotNull(line);
            using var doc = JsonDocument.Parse(line);
            var root = doc.RootElement;

            Assert.AreEqual("is_connected", root.GetProperty("tool").GetString());
            Assert.IsFalse(root.TryGetProperty("params", out _),
                "params key should be absent when parameters is null");
        }
        finally
        {
            client.Dispose();
            pipe.Dispose();
        }
    }

    [TestMethod]
    public async Task CallAsync_WithParams_SendsParamsObject()
    {
        var (client, sr, sw, pipe) = await CreatePairAsync(UniquePipe());
        try
        {
            var callTask = client.CallAsync("connect_access", new { database_path = @"C:\test.accdb" });

            var line = await sr.ReadLineAsync();
            await sw.WriteLineAsync("{\"success\":false,\"error\":\"not found\"}");
            await callTask;

            Assert.IsNotNull(line);
            using var doc = JsonDocument.Parse(line);
            var root = doc.RootElement;

            Assert.AreEqual("connect_access", root.GetProperty("tool").GetString());
            Assert.IsTrue(root.TryGetProperty("params", out var p));
            Assert.AreEqual(@"C:\test.accdb", p.GetProperty("database_path").GetString());
        }
        finally
        {
            client.Dispose();
            pipe.Dispose();
        }
    }

    [TestMethod]
    public async Task CallAsync_NullableFieldInParams_OmitsNullField()
    {
        // Simulates run_procedure with arguments_json = null
        var (client, sr, sw, pipe) = await CreatePairAsync(UniquePipe());
        try
        {
            var callTask = client.CallAsync("run_procedure",
                new { procedure_name = "MyProc", arguments_json = (string?)null });

            var line = await sr.ReadLineAsync();
            await sw.WriteLineAsync("{\"success\":true,\"value\":null}");
            await callTask;

            Assert.IsNotNull(line);
            using var doc = JsonDocument.Parse(line);
            var parms = doc.RootElement.GetProperty("params");

            Assert.AreEqual("MyProc", parms.GetProperty("procedure_name").GetString());
            Assert.IsFalse(parms.TryGetProperty("arguments_json", out _),
                "null optional field must be omitted from JSON");
        }
        finally
        {
            client.Dispose();
            pipe.Dispose();
        }
    }

    [TestMethod]
    public async Task CallAsync_IntegerParam_SerializedAsNumber()
    {
        // Verifies command_id / object_type arrive as JSON numbers, not strings
        var (client, sr, sw, pipe) = await CreatePairAsync(UniquePipe());
        try
        {
            var callTask = client.CallAsync("run_command", new { command_id = 42 });

            var line = await sr.ReadLineAsync();
            await sw.WriteLineAsync("{\"success\":false,\"error\":\"bad command\"}");
            await callTask;

            Assert.IsNotNull(line);
            using var doc = JsonDocument.Parse(line);
            var cmdId = doc.RootElement.GetProperty("params").GetProperty("command_id");

            Assert.AreEqual(JsonValueKind.Number, cmdId.ValueKind,
                "command_id must be a JSON number, not a string");
            Assert.AreEqual(42, cmdId.GetInt32());
        }
        finally
        {
            client.Dispose();
            pipe.Dispose();
        }
    }

    [TestMethod]
    public async Task CallAsync_ReturnsServerResponseVerbatim()
    {
        var (client, sr, sw, pipe) = await CreatePairAsync(UniquePipe());
        try
        {
            var expected = "{\"success\":true,\"forms\":[{\"name\":\"Form1\"}]}";
            var callTask = client.CallAsync("get_forms");

            await sr.ReadLineAsync();
            await sw.WriteLineAsync(expected);
            var result = await callTask;

            Assert.AreEqual(expected, result);
        }
        finally
        {
            client.Dispose();
            pipe.Dispose();
        }
    }

    // ── Concurrency ───────────────────────────────────────────────────────────

    [TestMethod]
    public async Task CallAsync_ConcurrentCalls_AreSerializedNotInterleaved()
    {
        // Two calls fired concurrently; the server must see them arrive one at a time
        // (not interleaved), because PipeChannel serializes with a SemaphoreSlim.
        var (client, sr, sw, pipe) = await CreatePairAsync(UniquePipe());
        try
        {
            var received = new List<string>();

            // Server loop: echo back each received request as-is
            var serverLoop = Task.Run(async () =>
            {
                for (int i = 0; i < 2; i++)
                {
                    var line = await sr.ReadLineAsync();
                    received.Add(line ?? "");
                    await sw.WriteLineAsync($"{{\"seq\":{i}}}");
                }
            });

            var t1 = client.CallAsync("get_forms");
            var t2 = client.CallAsync("get_reports");
            await Task.WhenAll(t1, t2, serverLoop);

            // Both requests arrived intact (no interleaving / partial writes)
            Assert.AreEqual(2, received.Count);
            foreach (var r in received)
            {
                using var doc = JsonDocument.Parse(r);
                Assert.IsTrue(doc.RootElement.TryGetProperty("tool", out _),
                    $"Received malformed request: {r}");
            }
        }
        finally
        {
            client.Dispose();
            pipe.Dispose();
        }
    }

    // ── Timeout / error handling ──────────────────────────────────────────────

    [TestMethod]
    public void Constructor_Timeout_ThrowsWhenNoPipeServer()
    {
        // No server listening → Connect should time out
        Assert.ThrowsExactly<TimeoutException>(() =>
        {
            using var _ = new PipeChannel("no-server-listening-xyzzy", connectTimeoutMs: 200);
        });
    }
}
