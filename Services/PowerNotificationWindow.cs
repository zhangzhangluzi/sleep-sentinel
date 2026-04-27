using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Diagnostics;

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
            try
            {
                var setting = Marshal.PtrToStructure<NativeMethods.PowerBroadcastSettingData>(m.LParam);
                if (setting.PowerSetting == NativeMethods.GuidLidSwitchStateChange)
                {
                    var handler = LidStateChanged;
                    try
                    {
                        handler?.Invoke(this, setting.Data != 0);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"LidState 监听回调失败：{ex.Message}");
                    }
                }
            }
            catch (ArgumentException ex)
            {
                Debug.WriteLine($"处理 LidState 通知失败：{ex.Message}");
            }
            catch (InvalidOperationException ex)
            {
                Debug.WriteLine($"处理 LidState 通知失败：{ex.Message}");
            }
        }

        base.WndProc(ref m);
    }
}
