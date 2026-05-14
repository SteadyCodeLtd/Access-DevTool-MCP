using System.Diagnostics;
using System.Text.Json;
using AccessAceMcp;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AccessAceMcpTests;

// End-to-end integration tests: .NET 10 PipeChannel -> Access-DevTool-Agent.exe -> AccessComService -> Access COM.
// These require Microsoft Access and the NorthwindAccess test database.
// Tests skip gracefully (Inconclusive) when the environment is not set up.
//
// Focus: the pipe layer's serialization, response shapes, and data flow —
// not a re-run of AccessComServiceTests. Each test asserts something about
// the JSON envelope or the round-trip that only goes through the pipe.

[TestClass]
[TestCategory("Integration")]
[DoNotParallelize]
public class WorkerPipeIntegrationTests
{
    private const string TestDb = @"C:\GitHub\NorthwindAccess\AccessMcp.accdb";

    // ── Shared worker (Access COM startup is slow; reuse across tests) ────────

    private static WorkerFixture? _shared;

    [ClassInitialize]
    public static void ClassInit(TestContext _)
    {
        if (!File.Exists(TestDb))
        {
            // Mark all tests in this class as inconclusive at setup time
            return;
        }

        _shared = WorkerFixture.Start(FindWorkerExe(), TestDb);
    }

    [ClassCleanup]
    public static async Task ClassCleanup()
    {
        if (_shared != null)
            await _shared.DisposeAsync();
    }

    private static PipeChannel Channel
    {
        get
        {
            if (!File.Exists(TestDb))
                Assert.Inconclusive($"Test database not found: {TestDb}");
            if (_shared == null)
                Assert.Inconclusive("Worker fixture not initialized (ClassInit skipped).");
            return _shared!.Channel;
        }
    }

    [TestInitialize]
    public async Task TestInit()
    {
        var r = AssertSuccess(await Channel.CallAsync("connect_access", new { database_path = TestDb }));
        Assert.IsTrue(r.GetProperty("connected").GetBoolean(), "Worker should be connected before each integration test");
    }

    // ── Worker helpers ────────────────────────────────────────────────────────

    private static string FindWorkerExe()
    {
        var dir = AppContext.BaseDirectory;
        var candidates = new[]
        {
            Path.GetFullPath(Path.Combine(dir, "..", "..", "..", "..", "Access-DevTool-Agent", "bin", "Debug", "Access-DevTool-Agent.exe")),
            Path.GetFullPath(Path.Combine(dir, "..", "..", "..", "..", "Access-DevTool-Agent", "bin", "Release", "Access-DevTool-Agent.exe")),
            Path.Combine(dir, "Access-DevTool-Agent.exe"),
        };

        var found = candidates.FirstOrDefault(File.Exists);
        if (found == null)
            Assert.Inconclusive("Access-DevTool-Agent.exe not found. Build Access-DevTool-Agent first.");
        return found!;
    }

    private sealed class WorkerFixture : IAsyncDisposable
    {
        public PipeChannel Channel { get; }
        private readonly Process _process;

        private WorkerFixture(Process p, PipeChannel c) { _process = p; Channel = c; }

        public static WorkerFixture Start(string workerExe, string? dbPath = null)
        {
            var pipeName = $"integ-test-{Guid.NewGuid():N}";
            var arguments = $"--pipe {pipeName}";
            if (dbPath != null) arguments += $" \"{dbPath}\"";

            var process = Process.Start(new ProcessStartInfo(workerExe)
            {
                Arguments             = arguments,
                UseShellExecute       = false,
                RedirectStandardError = true,
                CreateNoWindow        = true,
            }) ?? throw new InvalidOperationException("Failed to start worker");

            process.ErrorDataReceived += (_, e) =>
            {
                if (e.Data is { } line) Console.Error.WriteLine($"[worker-stderr] {line}");
            };
            process.BeginErrorReadLine();

            var channel = new PipeChannel(pipeName, connectTimeoutMs: 30_000);
            return new WorkerFixture(process, channel);
        }

        public async ValueTask DisposeAsync()
        {
            Channel.Dispose();
            try { if (!_process.HasExited) _process.Kill(entireProcessTree: true); }
            catch { /* best-effort */ }
            await _process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10)).ContinueWith(_ => { });
            _process.Dispose();
        }
    }

    private static JsonElement Parse(string json) =>
        JsonDocument.Parse(json).RootElement.Clone();

    private static JsonElement AssertSuccess(string json)
    {
        var el = Parse(json);
        if (!el.GetProperty("success").GetBoolean())
        {
            Assert.Inconclusive($"Integration environment returned non-success: {json}");
        }

        Assert.IsTrue(el.GetProperty("success").GetBoolean(),
            $"Expected success:true but got: {json}");
        return el;
    }

    // ── Auto-connect via CLI argument ─────────────────────────────────────────

    [TestMethod]
    public async Task AutoConnect_OnStartup_IsConnectedReturnsTrue()
    {
        // Worker was started with TestDb as arg in ClassInit — should auto-connect
        var r = AssertSuccess(await Channel.CallAsync("is_connected"));
        Assert.IsTrue(r.GetProperty("connected").GetBoolean(),
            "Worker should be connected after auto-connect via CLI argument");
    }

    // ── Response envelope shapes ──────────────────────────────────────────────

    [TestMethod]
    public async Task GetForms_ResponseHasFormsArray()
    {
        var r = AssertSuccess(await Channel.CallAsync("get_forms"));
        Assert.IsTrue(r.TryGetProperty("forms", out var forms),
            "Response must contain a 'forms' key");
        Assert.AreEqual(JsonValueKind.Array, forms.ValueKind);
        Assert.IsTrue(forms.GetArrayLength() > 0, "Expected at least one form");

        var first = forms[0];
        Assert.IsTrue(first.TryGetProperty("Name", out var name));
        Assert.IsFalse(string.IsNullOrWhiteSpace(name.GetString()));
    }

    [TestMethod]
    public async Task GetReports_ResponseHasReportsArray()
    {
        var r = AssertSuccess(await Channel.CallAsync("get_reports"));
        Assert.IsTrue(r.TryGetProperty("reports", out var reports));
        Assert.AreEqual(JsonValueKind.Array, reports.ValueKind);
        Assert.IsTrue(reports.GetArrayLength() > 0);
    }

    [TestMethod]
    public async Task GetQueries_ResponseHasQueriesWithSql()
    {
        var r = AssertSuccess(await Channel.CallAsync("get_queries"));
        Assert.IsTrue(r.TryGetProperty("queries", out var queries));

        var list = queries.EnumerateArray().ToList();
        Assert.IsTrue(list.Count > 0);
        Assert.IsTrue(list.Any(q =>
            q.TryGetProperty("Sql", out var sql) &&
            !string.IsNullOrWhiteSpace(sql.GetString())),
            "At least one query must have non-empty SQL");
    }

    [TestMethod]
    public async Task GetModules_ResponseHasModulesArray()
    {
        var r = AssertSuccess(await Channel.CallAsync("get_modules"));
        Assert.IsTrue(r.TryGetProperty("modules", out var modules));
        Assert.AreEqual(JsonValueKind.Array, modules.ValueKind);
        Assert.IsTrue(modules.GetArrayLength() > 0);
    }

    [TestMethod]
    public async Task GetApplicationInfo_ResponseHasExpectedKeys()
    {
        var r = AssertSuccess(await Channel.CallAsync("get_application_info"));
        Assert.IsTrue(r.TryGetProperty("info", out var info));
        Assert.IsTrue(info.TryGetProperty("Name", out _), "info must contain Name");
        Assert.IsTrue(info.TryGetProperty("Version", out _), "info must contain Version");
        Assert.IsTrue(info.TryGetProperty("CurrentDb", out _), "info must contain CurrentDb");
    }

    // ── Round-trip through the pipe ───────────────────────────────────────────

    [TestMethod]
    public async Task EvalExpression_ReturnsCorrectValue()
    {
        var r = AssertSuccess(await Channel.CallAsync("eval_expression", new { expression = "2+3" }));
        Assert.AreEqual("5", r.GetProperty("value").GetString());
    }

    [TestMethod]
    public async Task GetVbaCode_ThenSetVbaCode_RoundTripsViaChannel()
    {
        // Find the first available standard module across all projects
        var projectsResult = AssertSuccess(await Channel.CallAsync("get_vba_projects"));
        var projects = projectsResult.GetProperty("projects").EnumerateArray().ToList();

        if (projects.Count == 0)
            Assert.Inconclusive("No VBA projects found");

        string? moduleName = null;
        foreach (var project in projects)
        {
            if (!project.TryGetProperty("Components", out var components) || components.ValueKind != JsonValueKind.Array)
                continue;

            var stdModule = components.EnumerateArray()
                .FirstOrDefault(c => c.TryGetProperty("Type", out var t) && t.GetString() == "StandardModule");

            if (stdModule.ValueKind != JsonValueKind.Undefined)
            {
                moduleName = stdModule.GetProperty("Name").GetString();
                if (!string.IsNullOrWhiteSpace(moduleName))
                    break;
            }
        }

        if (string.IsNullOrWhiteSpace(moduleName))
            Assert.Inconclusive("No StandardModule found in VBA projects");

        // Get original code
        var codeResult = AssertSuccess(await Channel.CallAsync("get_vba_code", new { module_name = moduleName }));
        var original = codeResult.GetProperty("code").GetString()!;

        // Write modified code, verify it was written
        const string marker = "'-- PIPE-TEST-MARKER --\r\n";
        var writeResult = AssertSuccess(await Channel.CallAsync("set_vba_code",
            new { module_name = moduleName, code = marker + original }));
        Assert.IsTrue(writeResult.TryGetProperty("message", out _));

        var readBack = AssertSuccess(await Channel.CallAsync("get_vba_code", new { module_name = moduleName }));
        StringAssert.StartsWith(readBack.GetProperty("code").GetString(), marker,
            "Modified code must contain the written marker");

        // Restore original
        await Channel.CallAsync("set_vba_code", new { module_name = moduleName, code = original });

        var restored = AssertSuccess(await Channel.CallAsync("get_vba_code", new { module_name = moduleName }));
        Assert.AreEqual(original, restored.GetProperty("code").GetString(), "Code must be fully restored");
    }

    [TestMethod]
    public async Task ExportFormToText_ResponseContainsFormDefinition()
    {
        var formsResult = AssertSuccess(await Channel.CallAsync("get_forms"));
        var firstName = formsResult.GetProperty("forms")[0].GetProperty("Name").GetString()!;

        var r = AssertSuccess(await Channel.CallAsync("export_form_to_text", new { form_name = firstName }));
        var formData = r.GetProperty("form_data").GetString() ?? "";
        StringAssert.Contains(formData, "Begin Form",
            "Exported form text must contain 'Begin Form' header");
    }

    [TestMethod]
    public async Task ExportReportToText_ResponseContainsReportDefinition()
    {
        var reportsResult = AssertSuccess(await Channel.CallAsync("get_reports"));
        var firstName = reportsResult.GetProperty("reports")[0].GetProperty("Name").GetString()!;

        var r = AssertSuccess(await Channel.CallAsync("export_report_to_text", new { report_name = firstName }));
        var reportData = r.GetProperty("report_data").GetString() ?? "";
        StringAssert.Contains(reportData, "Begin Report",
            "Exported report text must contain 'Begin Report' header");
    }

    // ── Error propagation through the pipe ───────────────────────────────────

    [TestMethod]
    public async Task GetVbaCode_UnknownModule_ReturnsErrorViaChannel()
    {
        var result = Parse(await Channel.CallAsync("get_vba_code",
            new { module_name = "__NoSuchModule__" }));

        Assert.IsFalse(result.GetProperty("success").GetBoolean());
        Assert.IsTrue(result.TryGetProperty("error", out _));
    }

    [TestMethod]
    public async Task DeleteObject_UnknownForm_ReturnsErrorViaChannel()
    {
        // object_type 2 = acForm; passing an integer (not a string) tests the updated ParseIntArg
        var result = Parse(await Channel.CallAsync("delete_object",
            new { object_type = 2, object_name = "__NoSuchForm__" }));

        Assert.IsFalse(result.GetProperty("success").GetBoolean());
        Assert.IsTrue(result.TryGetProperty("error", out var err));
        Assert.IsFalse(err.GetString()?.Contains("must be an integer") ?? false,
            "object_type integer must not produce a parse error");
    }
}
