using System.Diagnostics;
using System.Collections.Generic;

namespace SleepSentinel.Services;

public sealed class PowerCfgService
{
    private const int TimeoutMilliseconds = 3000;
    private const int PowerCfgRetryDelayMilliseconds = 120;
    private const int PowerCfgQueryRetryCount = 2;
    private static readonly TimeSpan QueryResultCacheDuration = TimeSpan.FromSeconds(5);
    private static readonly TimeSpan QueryFailureCacheDuration = TimeSpan.FromSeconds(10);
    private static readonly string PowerCfgPath = ResolvePowerCfgPath();
    private readonly FileLogger _logger;
    private readonly object _runSync = new();
    private readonly Dictionary<string, CachedResult> _queryResultCache = [];

    public PowerCfgService(FileLogger logger)
    {
        _logger = logger;
    }

    public string Run(string arguments, int timeoutMilliseconds = TimeoutMilliseconds)
    {
        var normalizedArguments = arguments?.Trim() ?? string.Empty;
        var now = DateTimeOffset.UtcNow;

        lock (_runSync)
        {
            if (!IsQueryCommand(normalizedArguments))
            {
                _queryResultCache.Clear();
                var writeResult = ExecutePowerCfg(normalizedArguments, timeoutMilliseconds);
                _queryResultCache.Clear();
                return writeResult;
            }

            PurgeExpiredCache(now);
            if (_queryResultCache.TryGetValue(normalizedArguments, out var cached))
            {
                if (cached.ExpiredAtUtc > now)
                {
                    return cached.Result;
                }

                _queryResultCache.Remove(normalizedArguments);
            }

            var attempts = Math.Max(1, PowerCfgQueryRetryCount);
            string lastResult = string.Empty;
            for (var attempt = 1; attempt <= attempts; attempt++)
            {
                lastResult = ExecutePowerCfg(normalizedArguments, timeoutMilliseconds);
                if (!IsPowerCfgTimeoutResult(lastResult))
                {
                    break;
                }

                if (attempt < attempts)
                {
                    _logger.Warn($"powercfg {normalizedArguments} 第 {attempt} 次超时，等待 {PowerCfgRetryDelayMilliseconds}ms 后重试。");
                    Thread.Sleep(PowerCfgRetryDelayMilliseconds);
                }
            }

            var cacheTtl = IsPowerCfgTimeoutResult(lastResult)
                ? QueryFailureCacheDuration
                : QueryResultCacheDuration;
            _queryResultCache[normalizedArguments] = new CachedResult(lastResult, now.Add(cacheTtl));
            return lastResult;
        }
    }

    private string ExecutePowerCfg(string arguments, int timeoutMilliseconds = TimeoutMilliseconds)
    {
        try
        {
            using var process = new Process();
            process.StartInfo.FileName = PowerCfgPath;
            process.StartInfo.Arguments = arguments;
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;
            process.Start();

            var outputTask = process.StandardOutput.ReadToEndAsync();
            var errorTask = process.StandardError.ReadToEndAsync();

            if (!process.WaitForExit(timeoutMilliseconds))
            {
                TryKillProcess(process);
                process.WaitForExit();
                var timeoutMessage = $"powercfg {arguments} 超时（超过 {timeoutMilliseconds / 1000} 秒）";
                _logger.Warn(timeoutMessage);
                return timeoutMessage;
            }

            Task.WaitAll([outputTask, errorTask], 1000);
            var output = outputTask.IsCompletedSuccessfully ? outputTask.Result : string.Empty;
            var error = errorTask.IsCompletedSuccessfully ? errorTask.Result : string.Empty;

            if (process.ExitCode == 0)
            {
                return string.IsNullOrWhiteSpace(output) ? string.Empty : output.Trim();
            }

            var message = string.IsNullOrWhiteSpace(error) ? output.Trim() : error.Trim();
            return string.IsNullOrWhiteSpace(message) ? $"powercfg {arguments} 失败，退出码 {process.ExitCode}" : message;
        }
        catch (Exception ex)
        {
            _logger.Error($"执行 powercfg {arguments} 失败：{ex.Message}");
            return $"powercfg {arguments} 调用失败：{ex.Message}";
        }
    }

    private static string ResolvePowerCfgPath()
    {
        var candidate = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.System),
            "powercfg.exe");
        return File.Exists(candidate) ? candidate : "powercfg";
    }

    private static bool IsQueryCommand(string arguments)
    {
        return arguments.Equals(string.Empty, StringComparison.Ordinal)
            || IsCommandToken(arguments, "/q")
            || IsCommandToken(arguments, "/getactivescheme")
            || IsCommandToken(arguments, "/requests")
            || arguments.Equals("/requestsoverride", StringComparison.OrdinalIgnoreCase)
            || IsCommandToken(arguments, "/waketimers")
            || IsCommandToken(arguments, "/lastwake")
            || IsCommandToken(arguments, "/sleepstudy");
    }

    private static bool IsCommandToken(string arguments, string command)
    {
        return arguments.Equals(command, StringComparison.OrdinalIgnoreCase)
            || arguments.StartsWith(command + " ", StringComparison.OrdinalIgnoreCase)
            || arguments.StartsWith(command + "\t", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsPowerCfgTimeoutResult(string result)
    {
        return result.Contains("超时（超过", StringComparison.Ordinal);
    }

    private void PurgeExpiredCache(DateTimeOffset now)
    {
        if (_queryResultCache.Count == 0)
        {
            return;
        }

        var expiredKeys = _queryResultCache
            .Where(kv => kv.Value.ExpiredAtUtc <= now)
            .Select(kv => kv.Key)
            .ToList();

        foreach (var key in expiredKeys)
        {
            _queryResultCache.Remove(key);
        }
    }

    private readonly record struct CachedResult(string Result, DateTimeOffset ExpiredAtUtc);

    private static void TryKillProcess(Process process)
    {
        try
        {
            process.Kill(true);
        }
        catch (InvalidOperationException)
        {
        }
    }
}

