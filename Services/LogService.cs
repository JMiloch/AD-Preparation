using System.Collections.Concurrent;
using SCCMAdPrep.Models;

namespace SCCMAdPrep.Services;

/// <summary>
/// Thread-safe log service using ConcurrentQueue
/// </summary>
public class LogService
{
    private readonly ConcurrentQueue<LogEntry> _queue = new();

    public event Action? LogUpdated;

    public void Log(string message, LogLevel level = LogLevel.Text)
    {
        _queue.Enqueue(new LogEntry
        {
            Timestamp = DateTime.Now,
            Level = level,
            Message = message
        });
        LogUpdated?.Invoke();
    }

    public void Ok(string message) => Log(message, LogLevel.Ok);
    public void Info(string message) => Log(message, LogLevel.Info);
    public void Warn(string message) => Log(message, LogLevel.Warn);
    public void Error(string message) => Log(message, LogLevel.Error);
    public void DryRun(string message) => Log(message, LogLevel.DryRun);
    public void Section(string message) => Log(message, LogLevel.Section);

    public bool TryDequeue(out LogEntry? entry)
    {
        return _queue.TryDequeue(out entry);
    }
}
