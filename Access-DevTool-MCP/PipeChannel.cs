using System.IO.Pipes;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace AccessAceMcp;

/// <summary>
/// Thread-safe named pipe client that serializes JSON requests to the .NET 4.8 worker
/// and returns the raw JSON result string.
/// </summary>
internal sealed class PipeChannel : IDisposable
{
    private static readonly JsonSerializerOptions _opts = new()
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private readonly NamedPipeClientStream _pipe;
    private readonly StreamReader _reader;
    private readonly StreamWriter _writer;
    private readonly SemaphoreSlim _lock = new(1, 1);

    public PipeChannel(string pipeName, int connectTimeoutMs = 30_000)
    {
        _pipe = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut, PipeOptions.None);
        _pipe.Connect(connectTimeoutMs);
        _reader = new StreamReader(_pipe, new UTF8Encoding(false), detectEncodingFromByteOrderMarks: false, bufferSize: 65536, leaveOpen: true);
        _writer = new StreamWriter(_pipe, new UTF8Encoding(false), bufferSize: 65536, leaveOpen: true) { AutoFlush = true };
    }

    /// <summary>
    /// Sends a tool call to the worker and returns the JSON result string.
    /// Null-valued properties in <paramref name="parameters"/> are omitted from the request.
    /// </summary>
    public async Task<string> CallAsync(string toolName, object? parameters = null, CancellationToken ct = default)
    {
        await _lock.WaitAsync(ct);
        try
        {
            var request = JsonSerializer.Serialize(
                new { tool = toolName, @params = parameters },
                _opts);

            await _writer.WriteLineAsync(request.AsMemory(), ct);

            var response = await _reader.ReadLineAsync(ct)
                ?? throw new InvalidOperationException("Worker pipe disconnected unexpectedly.");

            return response;
        }
        finally
        {
            _lock.Release();
        }
    }

    public void Dispose()
    {
        try { _writer.Dispose(); } catch (ObjectDisposedException) { } catch (IOException) { }
        try { _reader.Dispose(); } catch (ObjectDisposedException) { } catch (IOException) { }
        try { _pipe.Dispose(); } catch (ObjectDisposedException) { } catch (IOException) { }
        _lock.Dispose();
    }
}
