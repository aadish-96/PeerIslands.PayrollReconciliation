namespace PayrollReconciliation.Services;

public class Logger
{
    private readonly string _logPath;
    private readonly object _lock = new();

    public Logger(string logPath)
    {
        _logPath = logPath;
        var dir = Path.GetDirectoryName(logPath);

        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);
        
        WriteRaw($"[{Ts()}] [INFO ] ===== Payroll Reconciliation Log Started =====");
    }

    public void Info(string message) => WriteRaw($"[{Ts()}] [INFO ] {message}");

    public void Warn(string message) => WriteRaw($"[{Ts()}] [WARN ] {message}");
    
    public void Error(string message) => WriteRaw($"[{Ts()}] [ERROR] {message}");

    private void WriteRaw(string line)
    {
        Console.WriteLine(line);

        lock (_lock)
            File.AppendAllText(_logPath, line + Environment.NewLine);
    }

    private static string Ts() => DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
}
