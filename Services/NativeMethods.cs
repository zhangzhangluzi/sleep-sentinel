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
    internal const int DeviceNotifyWindowHandle = 0x00000000;
    internal const int PbtPowerSettingChange = 0x8013;
    internal const int WmPowerBroadcast = 0x0218;
    internal static readonly Guid GuidLidSwitchStateChange = new("BA3E0F4D-B817-4094-A2D1-D56379E6A0F3");

    [StructLayout(LayoutKind.Sequential)]
    internal struct LastInputInfo
    {
        public uint cbSize;
        public uint dwTime;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct PowerBroadcastSettingData
    {
        public Guid PowerSetting;
        public uint DataLength;
        public int Data;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    internal static extern ExecutionState SetThreadExecutionState(ExecutionState esFlags);

    [DllImport("kernel32.dll")]
    internal static extern uint GetTickCount();

    [DllImport("user32.dll", SetLastError = false)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool GetLastInputInfo(ref LastInputInfo plii);

    [DllImport("user32.dll", SetLastError = true)]
    internal static extern IntPtr RegisterPowerSettingNotification(IntPtr hRecipient, ref Guid powerSettingGuid, int flags);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool UnregisterPowerSettingNotification(IntPtr handle);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool LockWorkStation();
}
