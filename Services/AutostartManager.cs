using System.Diagnostics;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace SleepSentinel.Services;

public static class AutostartManager
{
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string ValueName = "SleepSentinel";
    private const string ElevatedTaskName = "SleepSentinel Elevated Autostart";
    private const int ProcessTimeoutMilliseconds = 10000;
    private static readonly string WindowsPowerShellPath = ResolveWindowsPowerShellPath();
    private static readonly Regex CliXmlEnvelopeRegex = new(@"#<\s*CLIXML\s*", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex CliXmlPayloadRegex = new(@"<Objs[\s\S]*?</Objs>", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public enum AutostartMode
    {
        Disabled,
        RunKey,
        ElevatedScheduledTask
    }

    public readonly record struct AutostartStatus(AutostartMode Mode, bool MatchesDesiredConfiguration, string Summary, bool VerificationFailed = false);

    public static AutostartStatus QueryStatus(bool enabled, bool requireElevated)
    {
        var runKeyEnabled = IsRunKeyEnabled();
        var scheduledTask = QueryScheduledTask();
        if (scheduledTask.QueryFailed)
        {
            return BuildScheduledTaskQueryFailureStatus(enabled, requireElevated, runKeyEnabled, scheduledTask.QueryError);
        }

        var elevatedTaskEnabled = scheduledTask.Exists
            && scheduledTask.MatchesExecutablePath
            && scheduledTask.RunLevel.Equals("Highest", StringComparison.OrdinalIgnoreCase);

        var actualMode = elevatedTaskEnabled
            ? AutostartMode.ElevatedScheduledTask
            : runKeyEnabled
                ? AutostartMode.RunKey
                : AutostartMode.Disabled;
        var hasConflict = runKeyEnabled && elevatedTaskEnabled;

        if (!enabled)
        {
            var matchesDesired = actualMode == AutostartMode.Disabled && !hasConflict;
            return new AutostartStatus(
                actualMode,
                matchesDesired,
                matchesDesired
                    ? "开机自启未启用"
                    : "开机自启存在残留配置，建议重新应用");
        }

        if (requireElevated)
        {
            var matchesDesired = elevatedTaskEnabled && !runKeyEnabled;
            return new AutostartStatus(
                actualMode,
                matchesDesired,
                matchesDesired
                    ? "开机自启将通过计划任务以最高权限运行"
                    : runKeyEnabled
                        ? "开机自启当前仍是普通权限启动，重启后高权限保护可能无法自动接管"
                        : scheduledTask.Exists && !scheduledTask.MatchesExecutablePath
                            ? "检测到同名计划任务，但目标程序不是当前版本，建议重新应用开机自启"
                            : "开机自启尚未配置为最高权限启动");
        }

        var desiredRunKeyOnly = runKeyEnabled && !elevatedTaskEnabled;
        return new AutostartStatus(
            actualMode,
            desiredRunKeyOnly,
            desiredRunKeyOnly
                ? "开机自启将按当前用户普通权限运行"
                : elevatedTaskEnabled
                    ? "开机自启当前仍是最高权限计划任务，建议重新应用切回普通模式"
                    : "开机自启尚未配置");
    }

    public static AutostartStatus EnsureConfigured(bool enabled, bool requireElevated)
    {
        var current = QueryStatus(enabled, requireElevated);
        ThrowIfVerificationFailed(current);
        if (current.MatchesDesiredConfiguration)
        {
            return current;
        }

        SetEnabled(enabled, requireElevated);
        var verified = QueryStatus(enabled, requireElevated);
        ThrowIfVerificationFailed(verified);
        if (!verified.MatchesDesiredConfiguration)
        {
            throw new InvalidOperationException(verified.Summary);
        }

        return verified;
    }

    private static AutostartStatus BuildScheduledTaskQueryFailureStatus(bool enabled, bool requireElevated, bool runKeyEnabled, string queryError)
    {
        var actualMode = runKeyEnabled ? AutostartMode.RunKey : AutostartMode.Disabled;

        if (!enabled)
        {
            return new AutostartStatus(
                actualMode,
                false,
                $"开机自启状态读取失败：{queryError}",
                VerificationFailed: true);
        }

        if (requireElevated)
        {
            return new AutostartStatus(
                actualMode,
                false,
                runKeyEnabled
                    ? $"开机自启当前仍是普通权限启动，且最高权限计划任务状态读取失败：{queryError}"
                    : $"最高权限开机自启状态读取失败：{queryError}",
                VerificationFailed: true);
        }

        return new AutostartStatus(
            actualMode,
            false,
            runKeyEnabled
                ? $"普通开机自启已检测到，但计划任务状态读取失败，无法确认是否有残留最高权限任务：{queryError}"
                : $"开机自启状态读取失败：{queryError}",
            VerificationFailed: true);
    }

    private static void ThrowIfVerificationFailed(AutostartStatus status)
    {
        if (status.VerificationFailed)
        {
            throw new InvalidOperationException(status.Summary);
        }
    }

    public static void SetEnabled(bool enabled, bool requireElevated)
    {
        if (!enabled)
        {
            RemoveRunKey();
            RemoveElevatedScheduledTask();
            return;
        }

        if (requireElevated)
        {
            RegisterElevatedScheduledTask();
            RemoveRunKey();
            return;
        }

        SetRunKeyEnabled();
        RemoveElevatedScheduledTask();
    }

    private static string BuildCommand(string executablePath)
    {
        return $"\"{executablePath}\"";
    }

    private static void SetRunKeyEnabled()
    {
        using var key = Registry.CurrentUser.CreateSubKey(RunKeyPath, true);
        key?.SetValue(ValueName, BuildCommand(Application.ExecutablePath));
    }

    private static void RemoveRunKey()
    {
        using var key = Registry.CurrentUser.CreateSubKey(RunKeyPath, true);
        key?.DeleteValue(ValueName, false);
    }

    private static bool IsRunKeyEnabled()
    {
        using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, false);
        var configuredCommand = key?.GetValue(ValueName) as string;
        return string.Equals(
            NormalizeExecutablePath(configuredCommand),
            NormalizeExecutablePath(BuildCommand(Application.ExecutablePath)),
            StringComparison.OrdinalIgnoreCase);
    }

    private static ScheduledTaskInfo QueryScheduledTask()
    {
        const string script = """
            $ProgressPreference = 'SilentlyContinue'
            $task = Get-ScheduledTask -TaskName 'SleepSentinel Elevated Autostart' -ErrorAction SilentlyContinue
            if ($null -eq $task) {
                return
            }

            $action = $task.Actions | Select-Object -First 1
            $execute = if ($null -ne $action) { [string]$action.Execute } else { '' }
            $arguments = if ($null -ne $action) { [string]$action.Arguments } else { '' }
            $runLevel = [string]$task.Principal.RunLevel
            "{0}`t{1}`t{2}" -f $runLevel, $execute, $arguments
            """;

        var output = RunPowerShellScript(script);
        if (string.IsNullOrWhiteSpace(output))
        {
            return default;
        }

        if (IsPowerShellFailure(output))
        {
            return new ScheduledTaskInfo(false, false, string.Empty, output);
        }

        var parts = output.Split('\t');
        if (parts.Length < 3)
        {
            return new ScheduledTaskInfo(false, false, string.Empty, $"计划任务输出无法识别：{output}");
        }

        var executePath = NormalizeExecutablePath(parts[1]);
        var currentExecutablePath = NormalizeExecutablePath(Application.ExecutablePath);
        return new ScheduledTaskInfo(
            true,
            string.Equals(executePath, currentExecutablePath, StringComparison.OrdinalIgnoreCase),
            parts[0],
            string.Empty);
    }

    private static void RegisterElevatedScheduledTask()
    {
        var executablePath = EscapePowerShellString(Application.ExecutablePath);
        var userId = EscapePowerShellString(WindowsIdentity.GetCurrent().Name);
        var script = $@"
$ErrorActionPreference = 'Stop'
$action = New-ScheduledTaskAction -Execute '{executablePath}'
$trigger = New-ScheduledTaskTrigger -AtLogOn -User '{userId}'
$principal = New-ScheduledTaskPrincipal -UserId '{userId}' -LogonType InteractiveToken -RunLevel Highest
Register-ScheduledTask -TaskName '{ElevatedTaskName}' -Action $action -Trigger $trigger -Principal $principal -Force | Out-Null
";
        var output = RunPowerShellScript(script);
        if (!string.IsNullOrWhiteSpace(output))
        {
            throw new InvalidOperationException($"创建最高权限开机自启计划任务失败：{output}");
        }
    }

    private static void RemoveElevatedScheduledTask()
    {
        var script = $@"
$ProgressPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'
if (Get-ScheduledTask -TaskName '{ElevatedTaskName}' -ErrorAction SilentlyContinue) {{
    Unregister-ScheduledTask -TaskName '{ElevatedTaskName}' -Confirm:$false | Out-Null
}}
";
        var output = RunPowerShellScript(script);
        if (!string.IsNullOrWhiteSpace(output))
        {
            throw new InvalidOperationException($"移除最高权限开机自启计划任务失败：{output}");
        }
    }

    private static string RunPowerShellScript(string script)
    {
        try
        {
            using var process = new Process();
            var effectiveScript = "$ProgressPreference = 'SilentlyContinue'" + Environment.NewLine + script;
            var encodedCommand = Convert.ToBase64String(Encoding.Unicode.GetBytes(effectiveScript));
            process.StartInfo.FileName = WindowsPowerShellPath;
            process.StartInfo.Arguments = $"-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand {encodedCommand}";
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;
            process.Start();

            var outputTask = process.StandardOutput.ReadToEndAsync();
            var errorTask = process.StandardError.ReadToEndAsync();

            if (!process.WaitForExit(ProcessTimeoutMilliseconds))
            {
                TryKillProcess(process);
                process.WaitForExit();
                return $"PowerShell 脚本超时（超过 {ProcessTimeoutMilliseconds / 1000} 秒）";
            }

            Task.WaitAll([outputTask, errorTask], 1000);
            var output = SanitizePowerShellStream(outputTask.IsCompletedSuccessfully ? outputTask.Result : string.Empty);
            var error = SanitizePowerShellStream(errorTask.IsCompletedSuccessfully ? errorTask.Result : string.Empty);

            if (process.ExitCode == 0)
            {
                return string.IsNullOrWhiteSpace(output) ? string.Empty : output.Trim();
            }

            var message = string.IsNullOrWhiteSpace(error) ? output.Trim() : error.Trim();
            return string.IsNullOrWhiteSpace(message)
                ? $"PowerShell 脚本失败，退出码 {process.ExitCode}"
                : message;
        }
        catch (Exception ex)
        {
            return $"PowerShell 调用失败：{ex.Message}";
        }
    }

    private static string EscapePowerShellString(string value)
    {
        return value.Replace("'", "''", StringComparison.Ordinal);
    }

    private static string SanitizePowerShellStream(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        var sanitized = CliXmlEnvelopeRegex.Replace(text, string.Empty);
        sanitized = CliXmlPayloadRegex.Replace(sanitized, string.Empty);
        return sanitized.Trim();
    }

    private static bool IsPowerShellFailure(string output)
    {
        return output.StartsWith("PowerShell ", StringComparison.OrdinalIgnoreCase);
    }

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

    private static string ResolveWindowsPowerShellPath()
    {
        var candidate = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.System),
            "WindowsPowerShell",
            "v1.0",
            "powershell.exe");
        return File.Exists(candidate) ? candidate : "powershell";
    }

    private static string NormalizeExecutablePath(string? command)
    {
        if (string.IsNullOrWhiteSpace(command))
        {
            return string.Empty;
        }

        var trimmed = command.Trim();
        string executablePath;

        if (trimmed.StartsWith('"'))
        {
            var closingQuoteIndex = trimmed.IndexOf('"', 1);
            executablePath = closingQuoteIndex > 1 ? trimmed[1..closingQuoteIndex] : trimmed.Trim('"');
        }
        else
        {
            var firstSpaceIndex = trimmed.IndexOf(' ');
            executablePath = firstSpaceIndex > 0 ? trimmed[..firstSpaceIndex] : trimmed;
        }

        executablePath = Environment.ExpandEnvironmentVariables(executablePath);

        try
        {
            return Path.GetFullPath(executablePath);
        }
        catch (ArgumentException)
        {
            return executablePath;
        }
        catch (NotSupportedException)
        {
            return executablePath;
        }
        catch (PathTooLongException)
        {
            return executablePath;
        }
        catch (System.Security.SecurityException)
        {
            return executablePath;
        }
    }

    private readonly record struct ScheduledTaskInfo(bool Exists, bool MatchesExecutablePath, string RunLevel, string QueryError)
    {
        public bool QueryFailed => !string.IsNullOrWhiteSpace(QueryError);
    }
}
