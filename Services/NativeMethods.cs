using System.Runtime.InteropServices;

namespace SleepSentinel.Services;

[Flags]
public enum ExecutionState : uint
{
    EsAwayModeRequired = 0x00000040,
    EsContinuous = 0x80000000,
    EsDisplayRequired = 0x00000002,
    EsSystemRequired = 0x00000001
}

internal static class NativeMethods
{
    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    internal static extern ExecutionState SetThreadExecutionState(ExecutionState esFlags);
}
