using System.ComponentModel;
using System.Runtime.InteropServices;

namespace SleepSentinel.Services;

internal sealed class PowerNotificationWindow : NativeWindow, IDisposable
{
    private IntPtr _lidNotificationHandle;
    private bool _disposed;

    public PowerNotificationWindow()
    {
        CreateHandle(new CreateParams
        {
            Caption = "SleepSentinel.PowerNotificationWindow"
        });

        var guid = NativeMethods.GuidLidSwitchStateChange;
        _lidNotificationHandle = NativeMethods.RegisterPowerSettingNotification(
            Handle,
            ref guid,
            NativeMethods.DeviceNotifyWindowHandle);

        if (_lidNotificationHandle == IntPtr.Zero)
        {
            DestroyHandle();
            throw new Win32Exception(Marshal.GetLastWin32Error(), "注册开盖状态通知失败。");
        }
    }

    public event EventHandler<bool>? LidStateChanged;

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        if (_lidNotificationHandle != IntPtr.Zero)
        {
            NativeMethods.UnregisterPowerSettingNotification(_lidNotificationHandle);
            _lidNotificationHandle = IntPtr.Zero;
        }

        if (Handle != IntPtr.Zero)
        {
            DestroyHandle();
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == NativeMethods.WmPowerBroadcast
            && m.WParam == (IntPtr)NativeMethods.PbtPowerSettingChange
            && m.LParam != IntPtr.Zero)
        {
            var setting = Marshal.PtrToStructure<NativeMethods.PowerBroadcastSettingData>(m.LParam);
            if (setting.PowerSetting == NativeMethods.GuidLidSwitchStateChange)
            {
                LidStateChanged?.Invoke(this, setting.Data != 0);
            }
        }

        base.WndProc(ref m);
    }
}
