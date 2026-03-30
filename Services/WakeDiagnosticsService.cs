using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Globalization;
using System.ServiceProcess;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using SleepSentinel.Models;

namespace SleepSentinel.Services;

public sealed class WakeDiagnosticsService
{
    private static readonly Regex TimestampRegex = new(
        @"(?<timestamp>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly string[] ManualIndicators =
    [
        "power button",
        "input keyboard",
        "input mouse",
        "键盘",
        "鼠标",
        "开盖",
        "lid"
    ];

    private static readonly string[] SoftwareIndicators =
    [
        "windows update",
        "update orchestrator",
        "maintenance scheduler",
        "pdc task client",
        "maintenance activator",
        "updateorchestrator",
        "windowsupdateclient"
    ];

    private readonly PowerCfgService _powerCfg;
    private readonly FileLogger _logger;

    public WakeDiagnosticsService(PowerCfgService powerCfg, FileLogger logger)
    {
        _powerCfg = powerCfg;
        _logger = logger;
    }

    public WakeDiagnosticSnapshot CollectSnapshot(DateTimeOffset? sinceUtc = null, bool includePowerRequests = true, bool includeSleepStudy = true)
    {
        var snapshot = new WakeDiagnosticSnapshot
        {
            GeneratedAtUtc = DateTimeOffset.UtcNow,
            LastWakeText = _powerCfg.Run("/lastwake"),
            WakeTimersText = _powerCfg.Run("/waketimers"),
            EventSummary = CollectEventSummary(sinceUtc ?? DateTimeOffset.UtcNow.AddHours(-12))
        };

        if (includePowerRequests)
        {
            snapshot.PowerRequestsText = _powerCfg.Run("/requests");
            snapshot.RequestOverridesText = _powerCfg.Run("/requestsoverride");
            var suggestions = RemoteWakeBlockCatalog.SuggestCustomEntries(snapshot.PowerRequestsText);
            snapshot.SuggestedRemoteWakeEntriesSummary = suggestions.Count == 0
                ? "未从当前 requests / 运行进程 / 服务中发现新的远控候选项"
                : $"建议加入自定义远控拦截：{string.Join("、", suggestions)}";
        }

        if (includeSleepStudy)
        {
            snapshot.SleepStudySummary = CollectSleepStudySummary();
        }

        return snapshot;
    }

    public IReadOnlyList<string> SuggestCustomRemoteWakeEntries(IEnumerable<string>? existingEntries = null)
    {
        var requestsText = _powerCfg.Run("/requests");
        return RemoteWakeBlockCatalog.SuggestCustomEntries(requestsText, existingEntries);
    }

    public string SummarizeEvidence(WakeDiagnosticSnapshot snapshot)
    {
        if (TrySummarizeSleepToHibernatePowerButtonResume(snapshot, out var deepHibernateSummary))
        {
            return deepHibernateSummary;
        }

        if (!string.IsNullOrWhiteSpace(snapshot.EventSummary))
        {
            var firstLine = snapshot.EventSummary
                .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries)
                .FirstOrDefault();

            if (!string.IsNullOrWhiteSpace(firstLine))
            {
                return firstLine.Trim();
            }
        }

        if (ContainsAny(snapshot.LastWakeText, ManualIndicators))
        {
            return "powercfg 指向人工输入/开盖。";
        }

        if (ContainsAny(snapshot.LastWakeText, SoftwareIndicators) || ContainsAny(snapshot.WakeTimersText, SoftwareIndicators))
        {
            return "powercfg 指向软件/定时器活动。";
        }

        return "暂未从 powercfg 和事件日志中提取到明确证据。";
    }

    private string CollectEventSummary(DateTimeOffset sinceUtc)
    {
        var lines = new List<string>();
        lines.AddRange(ReadSystemHighlights(sinceUtc));
        lines.AddRange(ReadWindowsUpdateHighlights(sinceUtc));

        if (lines.Count == 0)
        {
            return $"自 {sinceUtc.ToLocalTime():yyyy-MM-dd HH:mm:ss} 以来未提取到关键电源/更新事件。";
        }

        return string.Join(Environment.NewLine, lines
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(static line => line, StringComparer.OrdinalIgnoreCase)
            .TakeLast(12));
    }

    private IEnumerable<string> ReadSystemHighlights(DateTimeOffset sinceUtc)
    {
        var allowedProviders = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Microsoft-Windows-Kernel-Power",
            "Service Control Manager",
            "Microsoft-Windows-Kernel-Boot"
        };

        var allowedIds = new HashSet<int> { 41, 42, 107, 506, 507, 7040, 7045, 20 };
        return ReadEventLines("System", sinceUtc, allowedProviders, allowedIds, 16);
    }

    private IEnumerable<string> ReadWindowsUpdateHighlights(DateTimeOffset sinceUtc)
    {
        var allowedProviders = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Microsoft-Windows-WindowsUpdateClient"
        };

        var allowedIds = new HashSet<int> { 19, 20, 21, 43, 44 };
        return ReadEventLines("Microsoft-Windows-WindowsUpdateClient/Operational", sinceUtc, allowedProviders, allowedIds, 10);
    }

    private IEnumerable<string> ReadEventLines(string logName, DateTimeOffset sinceUtc, ISet<string> providers, ISet<int> ids, int limit)
    {
        var lines = new List<string>();

        try
        {
            var query = new EventLogQuery(logName, PathType.LogName)
            {
                ReverseDirection = true
            };

            using var reader = new EventLogReader(query);
            for (EventRecord? record = reader.ReadEvent(); record is not null && lines.Count < limit; record = reader.ReadEvent())
            {
                using (record)
                {
                    if (record.TimeCreated is not { } createdAt)
                    {
                        continue;
                    }

                    var timestamp = new DateTimeOffset(createdAt);
                    if (timestamp < sinceUtc)
                    {
                        break;
                    }

                    if (!providers.Contains(record.ProviderName ?? string.Empty) || !ids.Contains(record.Id))
                    {
                        continue;
                    }

                    var message = record.FormatDescription();
                    if (string.IsNullOrWhiteSpace(message))
                    {
                        message = $"{record.ProviderName} 事件 {record.Id}";
                    }

                    message = CollapseWhitespace(message);
                    lines.Add($"{timestamp.ToLocalTime():yyyy-MM-dd HH:mm:ss} [{record.ProviderName}#{record.Id}] {message}");
                }
            }
        }
        catch (EventLogNotFoundException)
        {
        }
        catch (EventLogException ex)
        {
            _logger.Warn($"读取事件日志 {logName} 失败：{ex.Message}");
        }
        catch (UnauthorizedAccessException ex)
        {
            _logger.Warn($"读取事件日志 {logName} 被拒绝：{ex.Message}");
        }

        return lines;
    }

    private string CollectSleepStudySummary()
    {
        var outputPath = Path.Combine(Path.GetTempPath(), $"sleepstudy-{DateTime.Now:yyyyMMdd-HHmmss}.xml");
        var result = _powerCfg.Run($"/sleepstudy /duration 1 /xml /output \"{outputPath}\"");
        if (!File.Exists(outputPath))
        {
            return string.IsNullOrWhiteSpace(result) ? "未能生成 SleepStudy XML。" : result;
        }

        try
        {
            var document = XDocument.Load(outputPath);
            var ns = document.Root?.Name.Namespace;
            if (ns is null)
            {
                return "SleepStudy XML 结构不完整。";
            }

            var sleepStates = document
                .Descendants(ns + "OsStateInstance")
                .Select(element => new
                {
                    Element = element,
                    Type = (string?)element.Attribute("Type") ?? string.Empty,
                    EntryReason = (string?)element.Attribute("EntryReason") ?? string.Empty,
                    ExitReason = (string?)element.Attribute("ExitReason") ?? string.Empty,
                    StartLocal = ParseLocalTimestamp((string?)element.Attribute("LocalTimestamp")),
                    EndLocal = ParseLocalTimestamp((string?)element.Attribute("ExitLocalTimestamp"))
                })
                .OrderByDescending(static item => item.StartLocal)
                .ToList();

            var latestSleepState = sleepStates
                .FirstOrDefault(static item =>
                    item.Type.Equals("Hibernate", StringComparison.OrdinalIgnoreCase)
                    && item.EntryReason.Contains("Hibernate from Sleep", StringComparison.OrdinalIgnoreCase))
                ?? sleepStates.FirstOrDefault(static item => item.Type.Equals("Sleep", StringComparison.OrdinalIgnoreCase))
                ?? sleepStates.FirstOrDefault(static item =>
                    item.Type.Equals("Screen Off", StringComparison.OrdinalIgnoreCase)
                    && !string.IsNullOrWhiteSpace(item.ExitReason)
                    && !item.ExitReason.Equals("Unknown", StringComparison.OrdinalIgnoreCase));

            if (latestSleepState is null)
            {
                return "SleepStudy 中未找到最近的低功耗会话。";
            }

            var builder = new StringBuilder();
            builder.Append($"最近 SleepStudy 会话（{latestSleepState.Type}）：{latestSleepState.StartLocal:yyyy-MM-dd HH:mm:ss} -> {latestSleepState.EndLocal:yyyy-MM-dd HH:mm:ss}，进入={latestSleepState.EntryReason}，退出={latestSleepState.ExitReason}");

            if (latestSleepState.Type.Equals("Hibernate", StringComparison.OrdinalIgnoreCase)
                && latestSleepState.EntryReason.Contains("Hibernate from Sleep", StringComparison.OrdinalIgnoreCase)
                && latestSleepState.ExitReason.Contains("Power Button", StringComparison.OrdinalIgnoreCase))
            {
                builder.Append($"{Environment.NewLine}判定：系统先从待机自动转入更深休眠，最终由电源键成功唤醒；若中间有过第一次唤醒动作，它大概率未形成有效恢复。");
            }

            var scenarioInstanceId = latestSleepState.Element
                .Descendants(ns + "OSStateRecord")
                .FirstOrDefault(static record => string.Equals((string?)record.Attribute("Name"), "Scenario Instance ID", StringComparison.OrdinalIgnoreCase))
                ?.Attribute("Value")
                ?.Value;

            if (!string.IsNullOrWhiteSpace(scenarioInstanceId))
            {
                var scenario = document
                    .Descendants(ns + "ScenarioInstance")
                    .FirstOrDefault(element => string.Equals((string?)element.Attribute("Id"), scenarioInstanceId, StringComparison.OrdinalIgnoreCase));

                if (scenario is not null)
                {
                    var enterReason = (string?)scenario.Attribute("CsEnterReason");
                    var exitReason = (string?)scenario.Attribute("CsExitReason");
                    var blockers = scenario
                        .Descendants(ns + "TopBlocker")
                        .Take(3)
                        .Select(static blocker =>
                        {
                            var name = (string?)blocker.Attribute("Name") ?? "未知";
                            var active = (string?)blocker.Attribute("SegmentActiveTimePercent") ?? "0";
                            return $"{name} {active}%";
                        })
                        .ToArray();

                    if (!string.IsNullOrWhiteSpace(enterReason) || !string.IsNullOrWhiteSpace(exitReason))
                    {
                        builder.Append($"{Environment.NewLine}最近场景：进入={enterReason}，退出={exitReason}");
                    }

                    if (blockers.Length > 0)
                    {
                        builder.Append($"{Environment.NewLine}主要活跃项：{string.Join("、", blockers)}");
                    }
                }
            }

            return builder.ToString();
        }
        catch (IOException ex)
        {
            _logger.Warn($"读取 SleepStudy XML 失败：{ex.Message}");
            return $"已生成 SleepStudy XML，但读取失败：{ex.Message}";
        }
        catch (System.Xml.XmlException ex)
        {
            _logger.Warn($"解析 SleepStudy XML 失败：{ex.Message}");
            return $"已生成 SleepStudy XML，但解析失败：{ex.Message}";
        }
        finally
        {
            try
            {
                File.Delete(outputPath);
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }
    }

    private static DateTimeOffset ParseLocalTimestamp(string? value)
    {
        if (DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var timestamp))
        {
            return timestamp;
        }

        return DateTimeOffset.MinValue;
    }

    private static string CollapseWhitespace(string value)
    {
        var builder = new StringBuilder(value.Length);
        var previousWasWhitespace = false;

        foreach (var ch in value)
        {
            if (char.IsWhiteSpace(ch))
            {
                if (previousWasWhitespace)
                {
                    continue;
                }

                builder.Append(' ');
                previousWasWhitespace = true;
            }
            else
            {
                builder.Append(ch);
                previousWasWhitespace = false;
            }
        }

        return builder.ToString().Trim();
    }

    private static bool ContainsAny(string value, IEnumerable<string> indicators)
    {
        return indicators.Any(indicator => value.Contains(indicator, StringComparison.OrdinalIgnoreCase));
    }

    private static bool TrySummarizeSleepToHibernatePowerButtonResume(WakeDiagnosticSnapshot snapshot, out string summary)
    {
        var hibernateEvidence = FindEvidenceLine(snapshot.EventSummary, "hibernate from sleep - fixed timeout")
            ?? FindEvidenceLine(snapshot.SleepStudySummary, "hibernate from sleep - fixed timeout");
        var powerButtonEvidence = FindEvidenceLine(snapshot.EventSummary, "power button")
            ?? FindEvidenceLine(snapshot.SleepStudySummary, "power button");

        if (hibernateEvidence is null || powerButtonEvidence is null)
        {
            summary = string.Empty;
            return false;
        }

        var hibernateTimestamp = ExtractEvidenceTimestamp(hibernateEvidence);
        var powerButtonTimestamp = ExtractEvidenceTimestamp(powerButtonEvidence);

        summary = hibernateTimestamp is not null && powerButtonTimestamp is not null
            ? $"检测到系统先在 {hibernateTimestamp} 从待机自动转入更深休眠，直到 {powerButtonTimestamp} 才由电源键成功唤醒；第一次唤醒大概率未形成有效恢复。"
            : "检测到系统先从待机自动转入更深休眠，最终由电源键成功唤醒；第一次唤醒大概率未形成有效恢复。";
        return true;
    }

    private static string? FindEvidenceLine(string text, string keyword)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        return text
            .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries)
            .FirstOrDefault(line => line.Contains(keyword, StringComparison.OrdinalIgnoreCase));
    }

    private static string? ExtractEvidenceTimestamp(string line)
    {
        if (string.IsNullOrWhiteSpace(line))
        {
            return null;
        }

        var match = TimestampRegex.Match(line);
        if (!match.Success)
        {
            return null;
        }

        return match.Groups["timestamp"].Value;
    }
}
